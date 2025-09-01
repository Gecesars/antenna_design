# ctk_patch_array_hfss.py
# GUI (CustomTkinter) + PyAEDT para criar e simular um patch com 4 portas (Driven Modal + Lumped Ports)
# Requisitos:
#   pip install customtkinter ansys-aedt-core
#   Ansys Electronics Desktop (HFSS) instalado/licenciado
# Observação:
#   - Unidades em mm
#   - Porto lumped é uma folha vertical (entre ground @ z=0 e trilha/patch @ z=sub_h),
#     com linha de integração explícita em Z (evita erro "Both endpoints of port lines must lie on the port").

import os
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, filedialog

try:
    import customtkinter as ctk
except Exception as e:
    print("Faltando módulo 'customtkinter'. Instale com: pip install customtkinter")
    raise

# --- PyAEDT ---
try:
    from ansys.aedt.core import Desktop, Hfss
except Exception as e:
    Desktop = None
    Hfss = None

APP_TITLE = "HFSS Patch Array • Driven Modal + Lumped Ports (CTk)"
DEFAULT_PROJECT = "patch_array"
DEFAULT_DESIGN = "patch_array"

# -----------------------
# Utilidades de Geometria
# -----------------------
def must(obj, msg):
    """Garante que um objeto foi criado (Object3d/str); levanta erro caso contrário."""
    if obj is None:
        raise RuntimeError(msg)
    # Alguns métodos retornam bool; trate como falha
    if isinstance(obj, bool):
        if not obj:
            raise RuntimeError(msg)
        # True sem objeto não ajuda
        raise RuntimeError(f"{msg} (retorno inesperado 'bool')")
    # Se for string (nome de boundary), considere ok
    return obj

def add_lumped_port_vertical(hfss, name, center_xy, along_axis="x",
                             trace_w=2.0, inset=1.0, height_z=1.524):
    """
    Cria um Lumped Port como folha vertical:
      along_axis='x' -> folha está em plano YZ (x fixo), útil para trilha ao longo de X.
      along_axis='y' -> folha está em plano XZ (y fixo), útil para trilha ao longo de Y.
    A linha de integração é definida dentro da folha, de z≈0 até z≈height_z.
    """
    cx, cy = float(center_xy[0]), float(center_xy[1])

    if along_axis.lower() == "x":
        y1 = cy - trace_w / 2.0
        y2 = cy + trace_w / 2.0
        x_fixed = cx - float(inset)
        pts = [
            [x_fixed, y1, 0.0],
            [x_fixed, y2, 0.0],
            [x_fixed, y2, height_z],
            [x_fixed, y1, height_z],
            [x_fixed, y1, 0.0],
        ]
        mid = (y1 + y2) / 2.0
        int_p1 = [x_fixed, mid, 0.001]
        int_p2 = [x_fixed, mid, height_z - 0.001]
    else:
        x1 = cx - trace_w / 2.0
        x2 = cx + trace_w / 2.0
        y_fixed = cy - float(inset)
        pts = [
            [x1, y_fixed, 0.0],
            [x2, y_fixed, 0.0],
            [x2, y_fixed, height_z],
            [x1, y_fixed, height_z],
            [x1, y_fixed, 0.0],
        ]
        mid = (x1 + x2) / 2.0
        int_p1 = [mid, y_fixed, 0.001]
        int_p2 = [mid, y_fixed, height_z - 0.001]

    port_sheet = must(
        hfss.modeler.create_polyline(
            pts, cover_surface=True, closed=True, name=f"{name}_PortSheet"
        ),
        f"Falha ao criar folha do porto {name}"
    )

    ok = hfss.lumped_port(
        assignment=port_sheet.name,
        reference=None,
        integration_line=[int_p1, int_p2],
        impedance=50.0,
        name=f"{name}_Lumped",
        renormalize=True
    )
    if ok is False:
        raise RuntimeError(f"Falha ao atribuir Lumped Port {name}")
    return port_sheet

# -----------------------
# Núcleo de Simulação HFSS
# -----------------------
class HFSSPatchRunner:
    def __init__(self, gui, params):
        self.gui = gui
        self.p = params
        self.d = None
        self.hfss = None

    # Logging seguro na GUI
    def log(self, msg):
        self.gui.append_log(msg)

    def start_desktop(self):
        if Desktop is None or Hfss is None:
            raise RuntimeError("PyAEDT não encontrado. Instale 'ansys-aedt-core' e tente novamente.")
        self.log("Iniciando Electronics Desktop...")
        self.d = Desktop(new_desktop=True, non_graphical=False, close_on_exit=bool(self.p['fechar_ao_sair']))
        self.hfss = Hfss(
            projectname=self.p['project'],
            designname=self.p['design'],
            solution_type="DrivenModal"
        )
        self.hfss.modeler.model_units = "mm"
        self.log("HFSS pronto.")

    def build_geometry(self):
        # Materiais
        er = float(self.p['er'])
        tand = float(self.p['tand'])
        if "Rogers RO4003C (tm)" not in self.hfss.materials.material_keys:
            self.hfss.materials.add_material("Rogers RO4003C (tm)")
        m = self.hfss.materials["Rogers RO4003C (tm)"]
        m.permittivity = er
        m.dielectric_loss_tangent = tand

        sub_w = float(self.p['sub_w'])
        sub_l = float(self.p['sub_l'])
        sub_h = float(self.p['sub_h'])
        patch_w = float(self.p['patch_w'])
        patch_l = float(self.p['patch_l'])
        trace_w = float(self.p['trace_w'])
        stub_len = float(self.p['stub_len'])

        self.log("Criando Substrate...")
        sub = must(
            self.hfss.modeler.create_box(
                [-sub_w/2.0, -sub_l/2.0, 0.0],
                [sub_w, sub_l, sub_h],
                name="Substrate",
                matname="Rogers RO4003C (tm)"
            ), "Falha ao criar Substrate"
        )

        self.log("Criando Ground (sheet @ z=0)...")
        ground = must(
            self.hfss.modeler.create_rectangle(
                cs_plane="XY",
                position=[-sub_w/2.0, -sub_l/2.0],
                dimension_list=[sub_w, sub_l],
                name="Ground",
                is_covered=True
            ), "Falha ao criar Ground"
        )

        self.log("Criando Patch (sheet @ z=sub_h) e trilhas...")
        patch = must(
            self.hfss.modeler.create_rectangle(
                cs_plane="XY",
                position=[-patch_w/2.0, -patch_l/2.0],
                dimension_list=[patch_w, patch_l],
                name="Patch",
                is_covered=True
            ), "Falha ao criar Patch"
        )
        self.hfss.modeler.move([patch], [0, 0, sub_h])

        # Trilhas
        left_stub = must(self.hfss.modeler.create_rectangle(
            cs_plane="XY",
            position=[-patch_w/2.0 - stub_len, -trace_w/2.0],
            dimension_list=[stub_len, trace_w],
            name="Stub_Left", is_covered=True
        ), "Falha Stub_Left")
        right_stub = must(self.hfss.modeler.create_rectangle(
            cs_plane="XY",
            position=[patch_w/2.0, -trace_w/2.0],
            dimension_list=[stub_len, trace_w],
            name="Stub_Right", is_covered=True
        ), "Falha Stub_Right")
        bottom_stub = must(self.hfss.modeler.create_rectangle(
            cs_plane="XY",
            position=[-trace_w/2.0, -patch_l/2.0 - stub_len],
            dimension_list=[trace_w, stub_len],
            name="Stub_Bottom", is_covered=True
        ), "Falha Stub_Bottom")
        top_stub = must(self.hfss.modeler.create_rectangle(
            cs_plane="XY",
            position=[-trace_w/2.0, patch_l/2.0],
            dimension_list=[trace_w, stub_len],
            name="Stub_Top", is_covered=True
        ), "Falha Stub_Top")

        self.hfss.modeler.move([left_stub, right_stub, bottom_stub, top_stub], [0, 0, sub_h])
        # Une o condutor principal
        self.hfss.modeler.unite([patch, left_stub, right_stub, bottom_stub, top_stub])

        # Guarda para outros métodos
        self.sub_h = sub_h
        self.trace_w = trace_w
        self.patch_w = patch_w
        self.patch_l = patch_l
        self.ground = ground
        self.patch = patch

    def ports_and_boundaries(self):
        inset = float(self.p['port_inset'])
        # 4 portas — centros ao redor do patch
        self.log("Criando Lumped Ports...")
        add_lumped_port_vertical(self.hfss, "P1", center_xy=[-self.patch_w/2.0, 0.0], along_axis="x",
                                 trace_w=self.trace_w, inset=inset, height_z=self.sub_h)
        add_lumped_port_vertical(self.hfss, "P2", center_xy=[ self.patch_w/2.0, 0.0], along_axis="x",
                                 trace_w=self.trace_w, inset=inset, height_z=self.sub_h)
        add_lumped_port_vertical(self.hfss, "P3", center_xy=[0.0, -self.patch_l/2.0], along_axis="y",
                                 trace_w=self.trace_w, inset=inset, height_z=self.sub_h)
        add_lumped_port_vertical(self.hfss, "P4", center_xy=[0.0,  self.patch_l/2.0], along_axis="y",
                                 trace_w=self.trace_w, inset=inset, height_z=self.sub_h)

        # Ground como Perfect E
        self.log("Atribuindo Perfect E ao Ground...")
        must(self.hfss.assign_perfecte_to_sheets(self.ground), "Falha ao aplicar PerfectE no Ground")

        # Região + Radiação
        pad_air = [
            float(self.p['air_xm']), float(self.p['air_xp']),
            float(self.p['air_ym']), float(self.p['air_yp']),
            float(self.p['air_zm']), float(self.p['air_zp']),
        ]
        self.log("Criando região de ar e Radiation...")
        region = must(self.hfss.modeler.create_region(pad_air, is_percentage=False), "Falha ao criar região de ar")
        must(self.hfss.assign_radiation_boundary_to_objects(region), "Falha ao aplicar Radiation")

    def setup_and_solve(self):
        f0 = float(self.p['f0'])
        f1 = float(self.p['fs'])
        f2 = float(self.p['fe'])
        df = float(self.p['fd'])

        self.log("Criando Setup e Sweep...")
        setup = self.hfss.create_setup("Setup1")
        setup.props["Frequency"] = f"{f0}GHz"
        setup.props["MaxDeltaS"] = 0.02
        setup.update()

        sw = setup.create_linear_step_sweep(
            name="Sweep1", unit="GHz",
            start_frequency=f1, stop_frequency=f2, step_size=df,
            sweep_type="Interpolating"
        )
        sw.props["SaveFields"] = False
        sw.update()

        # Mesh mais fina no patch
        try:
            edge_len = max(self.trace_w, 0.5) / 6.0
            self.hfss.mesh.assign_length_mesh([self.patch], maximum_length=f"{edge_len:.3f}mm")
        except Exception:
            pass

        self.log("Rodando análise...")
        self.hfss.analyze_setup("Setup1")
        self.log("Análise concluída.")

    def export_sparams(self):
        # salva S11.png e Touchstone
        try:
            self.log("Exportando S-parâmetros e S11.png...")
            tdms = os.path.join(self.hfss.working_directory, f"{self.p['design']}.s2p")
            self.hfss.save_touchstone()
            plot = self.hfss.post.create_report(expressions="dB(S(1,1))", report_category="S Parameter")
            img = os.path.join(self.hfss.working_directory, "S11.png")
            self.hfss.post.export_report_to_png(plot, img, orientation="landscape", width=1200, height=800)
            self.log(f"OK • Touchstone/S11 em: {self.hfss.working_directory}")
        except Exception as e:
            self.log(f"Aviso: não foi possível exportar gráficos. {e}")

    def run(self):
        try:
            self.gui.set_running(True)
            self.start_desktop()
            self.build_geometry()
            self.ports_and_boundaries()
            self.setup_and_solve()
            self.export_sparams()
            self.log("FINALIZADO ✓")
        except Exception as e:
            self.log(f"ERRO: {e}")
            messagebox.showerror("Erro na simulação", str(e))
        finally:
            self.gui.set_running(False)
            if self.d and self.p['fechar_ao_sair']:
                # Desktop fechará automaticamente por close_on_exit=True
                pass

# -----------------------
# GUI (CustomTkinter)
# -----------------------
class PatchGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.title(APP_TITLE)
        self.geometry("1080x720")
        self.minsize(980, 640)

        # Estado
        self.running = False

        # --------- Layout base ----------
        self.columnconfigure(0, weight=0)  # sidebar
        self.columnconfigure(1, weight=1)  # main
        self.rowconfigure(0, weight=1)

        # Sidebar – ações
        self.sidebar = ctk.CTkFrame(self, corner_radius=16)
        self.sidebar.grid(row=0, column=0, sticky="nsw", padx=14, pady=14)
        self.sidebar.grid_rowconfigure(10, weight=1)

        self.title_lbl = ctk.CTkLabel(self.sidebar, text=APP_TITLE, wraplength=260, justify="left",
                                      font=ctk.CTkFont(size=16, weight="bold"))
        self.title_lbl.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")

        self.btn_run = ctk.CTkButton(self.sidebar, text="▶ Criar e Simular", command=self.on_run, height=42)
        self.btn_run.grid(row=1, column=0, padx=12, pady=(6, 6), sticky="ew")

        self.btn_save = ctk.CTkButton(self.sidebar, text="💾 Salvar Projeto .aedt", command=self.on_save_project)
        self.btn_save.grid(row=2, column=0, padx=12, pady=6, sticky="ew")

        self.open_dir_btn = ctk.CTkButton(self.sidebar, text="📂 Abrir Pasta de Trabalho",
                                          command=self.on_open_workdir)
        self.open_dir_btn.grid(row=3, column=0, padx=12, pady=6, sticky="ew")

        self.chk_close = ctk.CTkCheckBox(self.sidebar, text="Fechar HFSS ao sair", onvalue=True, offvalue=False)
        self.chk_close.select()
        self.chk_close.grid(row=4, column=0, padx=12, pady=(6, 12), sticky="w")

        self.progress = ctk.CTkProgressBar(self.sidebar)
        self.progress.set(0)
        self.progress.grid(row=5, column=0, padx=12, pady=(6, 12), sticky="ew")

        self.status = ctk.CTkLabel(self.sidebar, text="Pronto", anchor="w")
        self.status.grid(row=6, column=0, padx=12, pady=(0, 12), sticky="ew")

        # --------- Área principal: abas ----------
        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=1, sticky="nsew", padx=(0, 14), pady=14)

        self.tab_proj = self.tabs.add(" Projeto ")
        self.tab_geom = self.tabs.add(" Geometria ")
        self.tab_ports = self.tabs.add(" Portas & Região ")
        self.tab_solve = self.tabs.add(" Setup & Sweep ")
        self.tab_log = self.tabs.add(" Logs ")

        # --- Projeto ---
        self._build_tab_project()

        # --- Geometria ---
        self._build_tab_geometry()

        # --- Portas & Região ---
        self._build_tab_ports_region()

        # --- Setup & Sweep ---
        self._build_tab_solve()

        # --- Logs ---
        self.log_box = ctk.CTkTextbox(self.tab_log, height=400)
        self.log_box.pack(fill="both", expand=True, padx=12, pady=12)

        self.append_log("Bem-vindo! Ajuste os parâmetros e clique em 'Criar e Simular'.")

    # ---------- Construção das abas ----------
    def _build_tab_project(self):
        f = ctk.CTkFrame(self.tab_proj)
        f.pack(fill="x", padx=12, pady=12)

        self.entry_project = self._labeled_entry(f, "Nome do projeto (.aedt):", DEFAULT_PROJECT)
        self.entry_design = self._labeled_entry(f, "Nome do design:", DEFAULT_DESIGN)

    def _build_tab_geometry(self):
        grid = ctk.CTkFrame(self.tab_geom)
        grid.pack(fill="both", expand=True, padx=12, pady=12)

        # Substrato
        sub_frame = self._section(grid, "Substrato (RO4003C)")
        self.entry_sub_w = self._labeled_entry(sub_frame, "Largura (mm):", "60.0")
        self.entry_sub_l = self._labeled_entry(sub_frame, "Comprimento (mm):", "60.0")
        self.entry_sub_h = self._labeled_entry(sub_frame, "Altura h (mm):", "1.524")
        self.entry_er = self._labeled_entry(sub_frame, "εr:", "3.55")
        self.entry_tand = self._labeled_entry(sub_frame, "tanδ:", "0.0027")

        # Patch
        patch_frame = self._section(grid, "Patch")
        self.entry_patch_w = self._labeled_entry(patch_frame, "Largura Wp (mm):", "28.0")
        self.entry_patch_l = self._labeled_entry(patch_frame, "Comprimento Lp (mm):", "22.0")

        # Trilhas
        trace_frame = self._section(grid, "Trilhas (microstrip)")
        self.entry_trace_w = self._labeled_entry(trace_frame, "Largura trilha (mm):", "2.0")
        self.entry_stub_len = self._labeled_entry(trace_frame, "Comprimento stub (mm):", "5.0")

    def _build_tab_ports_region(self):
        grid = ctk.CTkFrame(self.tab_ports)
        grid.pack(fill="both", expand=True, padx=12, pady=12)

        ports_frame = self._section(grid, "Lumped Ports (verticais, Driven Modal)")
        self.entry_port_inset = self._labeled_entry(ports_frame, "Inset do porto (mm):", "1.0")
        ctk.CTkLabel(ports_frame, text="Os 4 portos (P1..P4) são criados automaticamente.").pack(anchor="w", padx=6, pady=4)

        air_frame = self._section(grid, "Região de Ar (distâncias absolutas em mm)")
        self.entry_air_xm = self._labeled_entry(air_frame, "-X (mm):", "10.0")
        self.entry_air_xp = self._labeled_entry(air_frame, "+X (mm):", "10.0")
        self.entry_air_ym = self._labeled_entry(air_frame, "-Y (mm):", "10.0")
        self.entry_air_yp = self._labeled_entry(air_frame, "+Y (mm):", "10.0")
        self.entry_air_zm = self._labeled_entry(air_frame, "-Z (mm):", "20.0")
        self.entry_air_zp = self._labeled_entry(air_frame, "+Z (mm):", "20.0")

    def _build_tab_solve(self):
        grid = ctk.CTkFrame(self.tab_solve)
        grid.pack(fill="x", padx=12, pady=12)

        freq_frame = self._section(grid, "Frequências (GHz)")
        self.entry_f0 = self._labeled_entry(freq_frame, "Frequência do Setup:", "10.0")
        sweep = self._section(grid, "Sweep Linear/Interpolating (GHz)")
        self.entry_fs = self._labeled_entry(sweep, "Início:", "8.0")
        self.entry_fe = self._labeled_entry(sweep, "Fim:", "12.0")
        self.entry_fd = self._labeled_entry(sweep, "Passo:", "0.05")

    # ---------- Helpers de UI ----------
    def _section(self, master, title):
        frame = ctk.CTkFrame(master, corner_radius=14)
        frame.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8, pady=(8, 4))
        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack(fill="x", padx=8, pady=(0, 8))
        return inner

    def _labeled_entry(self, master, label, default=""):
        row = ctk.CTkFrame(master, fg_color="transparent")
        row.pack(fill="x", pady=4)
        ctk.CTkLabel(row, text=label, width=220, anchor="w").pack(side="left", padx=(0, 8))
        entry = ctk.CTkEntry(row)
        entry.insert(0, str(default))
        entry.pack(side="left", fill="x", expand=True)
        return entry

    def append_log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{time.strftime('%H:%M:%S')}  {text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def set_running(self, running: bool):
        self.running = running
        self.btn_run.configure(state=("disabled" if running else "normal"))
        self.btn_save.configure(state=("disabled" if running else "normal"))
        self.open_dir_btn.configure(state=("disabled" if running else "normal"))
        self.progress.set(0.3 if running else 0.0)
        self.status.configure(text=("Processando..." if running else "Pronto"))

    # ---------- Ações ----------
    def collect_params(self):
        try:
            p = {
                "project": self.entry_project.get().strip() or DEFAULT_PROJECT,
                "design": self.entry_design.get().strip() or DEFAULT_DESIGN,
                "sub_w": float(self.entry_sub_w.get()),
                "sub_l": float(self.entry_sub_l.get()),
                "sub_h": float(self.entry_sub_h.get()),
                "er": float(self.entry_er.get()),
                "tand": float(self.entry_tand.get()),
                "patch_w": float(self.entry_patch_w.get()),
                "patch_l": float(self.entry_patch_l.get()),
                "trace_w": float(self.entry_trace_w.get()),
                "stub_len": float(self.entry_stub_len.get()),
                "port_inset": float(self.entry_port_inset.get()),
                "air_xm": float(self.entry_air_xm.get()),
                "air_xp": float(self.entry_air_xp.get()),
                "air_ym": float(self.entry_air_ym.get()),
                "air_yp": float(self.entry_air_yp.get()),
                "air_zm": float(self.entry_air_zm.get()),
                "air_zp": float(self.entry_air_zp.get()),
                "f0": float(self.entry_f0.get()),
                "fs": float(self.entry_fs.get()),
                "fe": float(self.entry_fe.get()),
                "fd": float(self.entry_fd.get()),
                "fechar_ao_sair": bool(self.chk_close.get()),
            }
        except ValueError as e:
            raise RuntimeError(f"Parâmetro inválido: {e}")
        return p

    def on_run(self):
        if self.running:
            return
        try:
            params = self.collect_params()
        except Exception as e:
            messagebox.showerror("Parâmetros inválidos", str(e))
            return

        def _task():
            runner = HFSSPatchRunner(self, params)
            runner.run()

        threading.Thread(target=_task, daemon=True).start()

    def on_save_project(self):
        # salva a .aedt no local escolhido (o HFSS salvará ao encerrar)
        folder = filedialog.askdirectory(title="Escolha a pasta para salvar o projeto .aedt")
        if not folder:
            return
        # Apenas informa — o PyAEDT salva durante a sessão; após rodar, peça para salvar manualmente se preciso.
        messagebox.showinfo("Salvar Projeto", "Após a simulação, use 'File > Save' no HFSS para gravar na pasta escolhida.")

    def on_open_workdir(self):
        try:
            # Tenta abrir a pasta temp do usuário – melhor após criar a sessão.
            path = os.path.expanduser("~")
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception as e:
            messagebox.showwarning("Não foi possível abrir a pasta", str(e))


def main():
    app = PatchGUI()
    app.mainloop()

if __name__ == "__main__":
    main()
