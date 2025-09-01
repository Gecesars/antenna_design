# ctk_patch_array_hfss.py
# GUI (CustomTkinter) + PyAEDT para criar e simular um PATCH parametrizado
# 4 portas lumped (Driven Modal) em folhas verticais com linha de integração correta.
# > Requer: Ansys Electronics Desktop (HFSS) instalado/licenciado.

import os
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, filedialog

# -------- UI --------
try:
    import customtkinter as ctk
except Exception:
    print("Instale o CustomTkinter:  pip install customtkinter")
    raise

# -------- PyAEDT --------
try:
    from ansys.aedt.core import Desktop, Hfss
except Exception:
    Desktop = None
    Hfss = None

APP_TITLE = "HFSS Patch Paramétrico • Driven Modal + Lumped Ports (CTk)"
DEFAULT_PROJECT = "patch_array"
DEFAULT_DESIGN = "patch_array"

# =========================
# Utilidades de geometria
# =========================
def ok_or_raise(val, msg_fail):
    """Aceita objetos ou True; rejeita None/False."""
    if val is None or val is False:
        raise RuntimeError(msg_fail)
    return val

def mm_expr(value_float):
    """Converte float -> string com unidade mm."""
    return f"{float(value_float)}mm"

def add_lumped_port_vertical_param(
    hfss, name, axis, sub_h_var="sub_h", patch_w_var="patch_w",
    patch_l_var="patch_l", trace_w_var="trace_w", inset_var="port_inset"
):
    """
    Cria uma folha (polyline fechada) vertical para Lumped Port:
      - axis = 'x_left' | 'x_right' | 'y_bottom' | 'y_top'
      - usa apenas EXPRESSÕES PARAMÉTRICAS (strings) com mm
    Integração: reta interna ao plano da folha, de z~0 a z~sub_h.
    """

    # Coordenadas paramétricas em strings:
    z0 = "0mm"
    zh = sub_h_var
    eps = "0.001mm"
    tw2 = f"{trace_w_var}/2"

    if axis == "x_left":
        # x = -patch_w/2 - inset ; y in [-tw/2, +tw/2]
        xfix = f"(-{patch_w_var}/2 - {inset_var})"
        y1 = f"-{tw2}"
        y2 = f"{tw2}"
        pts = [
            [xfix, y1, z0],
            [xfix, y2, z0],
            [xfix, y2, zh],
            [xfix, y1, zh],
            [xfix, y1, z0],
        ]
        int_p1 = [xfix, "0mm", eps]
        int_p2 = [xfix, "0mm", f"{zh}-{eps}"]

    elif axis == "x_right":
        xfix = f"({patch_w_var}/2 + {inset_var})"
        y1 = f"-{tw2}"
        y2 = f"{tw2}"
        pts = [
            [xfix, y1, z0],
            [xfix, y2, z0],
            [xfix, y2, zh],
            [xfix, y1, zh],
            [xfix, y1, z0],
        ]
        int_p1 = [xfix, "0mm", eps]
        int_p2 = [xfix, "0mm", f"{zh}-{eps}"]

    elif axis == "y_bottom":
        # y = -patch_l/2 - inset ; x in [-tw/2, +tw/2]
        yfix = f"(-{patch_l_var}/2 - {inset_var})"
        x1 = f"-{tw2}"
        x2 = f"{tw2}"
        pts = [
            [x1, yfix, z0],
            [x2, yfix, z0],
            [x2, yfix, zh],
            [x1, yfix, zh],
            [x1, yfix, z0],
        ]
        int_p1 = ["0mm", yfix, eps]
        int_p2 = ["0mm", yfix, f"{zh}-{eps}"]

    elif axis == "y_top":
        yfix = f"({patch_l_var}/2 + {inset_var})"
        x1 = f"-{tw2}"
        x2 = f"{tw2}"
        pts = [
            [x1, yfix, z0],
            [x2, yfix, z0],
            [x2, yfix, zh],
            [x1, yfix, zh],
            [x1, yfix, z0],
        ]
        int_p1 = ["0mm", yfix, eps]
        int_p2 = ["0mm", yfix, f"{zh}-{eps}"]

    else:
        raise ValueError("axis inválido")

    sheet = ok_or_raise(
        hfss.modeler.create_polyline(pts, cover_surface=True, closed=True, name=f"{name}_PortSheet"),
        f"Falha ao criar {name}_PortSheet"
    )

    ok = hfss.lumped_port(
        assignment=sheet.name,
        reference=None,
        integration_line=[int_p1, int_p2],
        impedance=50.0,
        name=name,
        renormalize=True
    )
    if ok is False:
        raise RuntimeError(f"Falha ao atribuir Lumped Port {name}")
    return sheet

# =========================
# Núcleo HFSS
# =========================
class HFSSPatchRunner:
    def __init__(self, gui, params):
        self.gui = gui
        self.p = params
        self.d = None
        self.hfss = None

    def log(self, msg):
        self.gui.append_log(msg)

    # --------- Sessão ---------
    def start_desktop(self):
        if Desktop is None or Hfss is None:
            raise RuntimeError("PyAEDT não disponível. Instale: pip install ansys-aedt-core")
        self.log("Iniciando Ansys Electronics Desktop...")
        self.d = Desktop(new_desktop=True, non_graphical=False, close_on_exit=bool(self.p['fechar_ao_sair']))
        self.hfss = Hfss(
            projectname=self.p['project'],
            designname=self.p['design'],
            solution_type="DrivenModal"
        )
        self.hfss.modeler.model_units = "mm"
        self.log("HFSS aberto e pronto.")

    # --------- Variáveis (parametrização) ---------
    def declare_vars(self):
        H = self.hfss
        # Geometria
        H["sub_w"]     = mm_expr(self.p['sub_w'])
        H["sub_l"]     = mm_expr(self.p['sub_l'])
        H["sub_h"]     = mm_expr(self.p['sub_h'])
        H["patch_w"]   = mm_expr(self.p['patch_w'])
        H["patch_l"]   = mm_expr(self.p['patch_l'])
        H["trace_w"]   = mm_expr(self.p['trace_w'])
        H["stub_len"]  = mm_expr(self.p['stub_len'])
        H["port_inset"]= mm_expr(self.p['port_inset'])

        # Região de ar
        H["air_xm"] = mm_expr(self.p['air_xm'])
        H["air_xp"] = mm_expr(self.p['air_xp'])
        H["air_ym"] = mm_expr(self.p['air_ym'])
        H["air_yp"] = mm_expr(self.p['air_yp'])
        H["air_zm"] = mm_expr(self.p['air_zm'])
        H["air_zp"] = mm_expr(self.p['air_zp'])

        # Materiais
        H["er_sub"]   = str(float(self.p['er']))
        H["tand_sub"] = str(float(self.p['tand']))

    # --------- Materiais ---------
    def materials(self):
        H = self.hfss
        if "Rogers RO4003C (tm)" not in H.materials.material_keys:
            H.materials.add_material("Rogers RO4003C (tm)")
        mat = H.materials["Rogers RO4003C (tm)"]
        mat.permittivity = float(self.p['er'])
        mat.dielectric_loss_tangent = float(self.p['tand'])

    # --------- Geometria ---------
    def build_geometry(self):
        self.log("Declarando variáveis...")
        self.declare_vars()

        self.log("Configurando materiais...")
        self.materials()

        H = self.hfss
        # Substrato (box) paramétrico
        self.log("Criando Substrate (box)...")
        sub = ok_or_raise(
            H.modeler.create_box(
                # origem no centro negativo, em expressões:
                ["-sub_w/2", "-sub_l/2", "0mm"],
                ["sub_w", "sub_l", "sub_h"],
                name="Substrate",
                matname="Rogers RO4003C (tm)"
            ),
            "Falha ao criar Substrate"
        )

        # Ground (sheet XY @ z=0) — API atual requer orientation
        self.log("Criando Ground (sheet em z=0)...")
        ground = ok_or_raise(
            H.modeler.create_rectangle(
                origin=["-sub_w/2", "-sub_l/2"],
                sizes=["sub_w", "sub_l"],
                name="Ground",
                is_covered=True,
                orientation="XY"
            ),
            "Falha ao criar Ground"
        )

        # Patch (sheet XY) criado em z=0 e movido para z=sub_h
        self.log("Criando Patch (sheet @ z=sub_h) e trilhas...")
        patch = ok_or_raise(
            H.modeler.create_rectangle(
                origin=["-patch_w/2", "-patch_l/2"],
                sizes=["patch_w", "patch_l"],
                name="Patch",
                is_covered=True,
                orientation="XY"
            ),
            "Falha ao criar Patch"
        )
        H.modeler.move([patch], [ "0mm", "0mm", "sub_h" ])

        # Trilhas (sheets XY) — todas em z=sub_h
        left_stub = ok_or_raise(
            H.modeler.create_rectangle(
                origin=["-patch_w/2 - stub_len", "-trace_w/2"],
                sizes=["stub_len", "trace_w"],
                name="Stub_Left",
                is_covered=True,
                orientation="XY"
            ),
            "Falha Stub_Left"
        )
        right_stub = ok_or_raise(
            H.modeler.create_rectangle(
                origin=["patch_w/2", "-trace_w/2"],
                sizes=["stub_len", "trace_w"],
                name="Stub_Right",
                is_covered=True,
                orientation="XY"
            ),
            "Falha Stub_Right"
        )
        bottom_stub = ok_or_raise(
            H.modeler.create_rectangle(
                origin=["-trace_w/2", "-patch_l/2 - stub_len"],
                sizes=["trace_w", "stub_len"],
                name="Stub_Bottom",
                is_covered=True,
                orientation="XY"
            ),
            "Falha Stub_Bottom"
        )
        top_stub = ok_or_raise(
            H.modeler.create_rectangle(
                origin=["-trace_w/2", "patch_l/2"],
                sizes=["trace_w", "stub_len"],
                name="Stub_Top",
                is_covered=True,
                orientation="XY"
            ),
            "Falha Stub_Top"
        )
        H.modeler.move([left_stub, right_stub, bottom_stub, top_stub], ["0mm", "0mm", "sub_h"])
        H.modeler.unite([patch, left_stub, right_stub, bottom_stub, top_stub])

        # Guarda referências
        self.ground = ground
        self.patch = patch

    # --------- Portas e Boundaries ---------
    def ports_and_boundaries(self):
        H = self.hfss
        self.log("Criando Lumped Ports paramétricos (folhas verticais)...")
        add_lumped_port_vertical_param(H, "P1", "x_left")
        add_lumped_port_vertical_param(H, "P2", "x_right")
        add_lumped_port_vertical_param(H, "P3", "y_bottom")
        add_lumped_port_vertical_param(H, "P4", "y_top")

        self.log("Atribuindo PerfectE ao Ground...")
        H.assign_perfecte_to_sheets(self.ground)

        # Região de ar absoluta (mm). create_region nomeia como "Region"
        self.log("Criando região de ar (absoluta, mm) e boundary Radiation...")
        # Ordem: [-X, +X, -Y, +Y, -Z, +Z]
        region = H.modeler.create_region(
            ["air_xm", "air_xp", "air_ym", "air_yp", "air_zm", "air_zp"],
            is_percentage=False
        )
        # Independente do retorno, aplicamos Radiaton ao objeto "Region"
        H.assign_radiation_boundary_to_objects("Region")

    # --------- Setup & Solve ---------
    def setup_and_solve(self):
        H = self.hfss
        f0 = float(self.p['f0'])
        fs = float(self.p['fs'])
        fe = float(self.p['fe'])
        fd = float(self.p['fd'])

        self.log("Criando Setup 'Setup1' e Sweep 'Sweep1'...")
        setup = H.create_setup("Setup1")
        setup.props["Frequency"] = f"{f0}GHz"
        setup.props["MaxDeltaS"] = 0.02
        setup.update()

        sweep = setup.create_linear_step_sweep(
            name="Sweep1",
            unit="GHz",
            start_frequency=fs,
            stop_frequency=fe,
            step_size=fd,
            sweep_type="Interpolating"
        )
        sweep.update()

        # Malha fina no patch
        try:
            H.mesh.assign_length_mesh([self.patch], maximum_length="(trace_w/6)")
        except Exception:
            pass

        self.log("Rodando simulação...")
        H.analyze_setup("Setup1")
        self.log("Simulação concluída.")

    # --------- Export ---------
    def export(self):
        H = self.hfss
        try:
            self.log("Exportando Touchstone e gráfico S11...")
            H.save_touchstone()  # salva em working_directory
            rep = H.post.create_report("dB(S(1,1))", report_category="S Parameter")
            png = os.path.join(H.working_directory, "S11.png")
            H.post.export_report_to_png(rep, png, width=1400, height=900)
            self.log(f"OK • Arquivos: {H.working_directory}")
        except Exception as e:
            self.log(f"Aviso na exportação: {e}")

    # --------- Fluxo total ---------
    def run(self):
        try:
            self.gui.set_running(True)
            self.start_desktop()
            self.build_geometry()
            self.ports_and_boundaries()
            self.setup_and_solve()
            self.export()
            self.log("FINALIZADO ✓")
        except Exception as e:
            self.log(f"ERRO: {e}")
            messagebox.showerror("Erro na simulação", str(e))
        finally:
            self.gui.set_running(False)
            # close_on_exit já definido no Desktop

# =========================
# GUI (CustomTkinter)
# =========================
class PatchGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.title(APP_TITLE)
        self.geometry("1120x740")
        self.minsize(1000, 680)

        self.running = False

        # Layout base
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar = ctk.CTkFrame(self, corner_radius=16)
        self.sidebar.grid(row=0, column=0, sticky="nsw", padx=14, pady=14)
        self.sidebar.grid_rowconfigure(10, weight=1)

        ctk.CTkLabel(self.sidebar, text=APP_TITLE, wraplength=260, justify="left",
                     font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=12, pady=(12,6), sticky="w")

        self.btn_run = ctk.CTkButton(self.sidebar, text="▶ Criar e Simular", height=44, command=self.on_run)
        self.btn_run.grid(row=1, column=0, padx=12, pady=(6,6), sticky="ew")

        self.btn_save = ctk.CTkButton(self.sidebar, text="💾 Salvar Projeto .aedt", command=self.on_save_project)
        self.btn_save.grid(row=2, column=0, padx=12, pady=6, sticky="ew")

        self.btn_open = ctk.CTkButton(self.sidebar, text="📂 Abrir Pasta de Trabalho", command=self.on_open_workdir)
        self.btn_open.grid(row=3, column=0, padx=12, pady=6, sticky="ew")

        self.chk_close = ctk.CTkCheckBox(self.sidebar, text="Fechar HFSS ao sair", onvalue=True, offvalue=False)
        self.chk_close.select()
        self.chk_close.grid(row=4, column=0, padx=12, pady=(6,12), sticky="w")

        self.progress = ctk.CTkProgressBar(self.sidebar)
        self.progress.set(0)
        self.progress.grid(row=5, column=0, padx=12, pady=(6,12), sticky="ew")

        self.status = ctk.CTkLabel(self.sidebar, text="Pronto", anchor="w")
        self.status.grid(row=6, column=0, padx=12, pady=(0,12), sticky="ew")

        # Abas
        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=1, sticky="nsew", padx=(0,14), pady=14)

        self.tab_proj = self.tabs.add(" Projeto ")
        self.tab_geom = self.tabs.add(" Geometria ")
        self.tab_ports = self.tabs.add(" Portas & Região ")
        self.tab_solve = self.tabs.add(" Setup & Sweep ")
        self.tab_log = self.tabs.add(" Logs ")

        self._build_tab_project()
        self._build_tab_geometry()
        self._build_tab_ports_region()
        self._build_tab_solve()

        # Logs
        self.log_box = ctk.CTkTextbox(self.tab_log)
        self.log_box.pack(fill="both", expand=True, padx=12, pady=12)
        self.append_log("Bem-vindo! Ajuste os parâmetros e clique em 'Criar e Simular'.")

    # ---- Seções das abas ----
    def _build_tab_project(self):
        f = ctk.CTkFrame(self.tab_proj)
        f.pack(fill="x", padx=12, pady=12)
        self.entry_project = self._labeled_entry(f, "Nome do projeto (.aedt):", DEFAULT_PROJECT)
        self.entry_design  = self._labeled_entry(f, "Nome do design:", DEFAULT_DESIGN)

    def _build_tab_geometry(self):
        grid = ctk.CTkFrame(self.tab_geom)
        grid.pack(fill="both", expand=True, padx=12, pady=12)

        sub = self._section(grid, "Substrato (RO4003C) — Paramétrico")
        self.entry_sub_w = self._labeled_entry(sub, "Largura sub_w (mm):", "60.0")
        self.entry_sub_l = self._labeled_entry(sub, "Comprimento sub_l (mm):", "60.0")
        self.entry_sub_h = self._labeled_entry(sub, "Altura sub_h (mm):", "1.524")
        self.entry_er    = self._labeled_entry(sub, "εr:", "3.55")
        self.entry_tand  = self._labeled_entry(sub, "tanδ:", "0.0027")

        patch = self._section(grid, "Patch — Paramétrico")
        self.entry_patch_w = self._labeled_entry(patch, "Largura patch_w (mm):", "28.0")
        self.entry_patch_l = self._labeled_entry(patch, "Compr. patch_l (mm):", "22.0")

        trace = self._section(grid, "Trilhas (microstrip) — Paramétrico")
        self.entry_trace_w = self._labeled_entry(trace, "Largura trace_w (mm):", "2.0")
        self.entry_stub_len= self._labeled_entry(trace, "Compr. stub_len (mm):", "5.0")

    def _build_tab_ports_region(self):
        grid = ctk.CTkFrame(self.tab_ports)
        grid.pack(fill="both", expand=True, padx=12, pady=12)

        ports = self._section(grid, "Lumped Ports (Driven Modal)")
        self.entry_port_inset = self._labeled_entry(ports, "port_inset (mm):", "1.0")
        ctk.CTkLabel(ports, text="P1/P2 nas laterais (X), P3/P4 topo/base (Y).").pack(anchor="w", padx=6, pady=4)

        reg = self._section(grid, "Região de Ar (absoluta, mm)")
        self.entry_air_xm = self._labeled_entry(reg, "air_xm (−X, mm):", "10.0")
        self.entry_air_xp = self._labeled_entry(reg, "air_xp (+X, mm):", "10.0")
        self.entry_air_ym = self._labeled_entry(reg, "air_ym (−Y, mm):", "10.0")
        self.entry_air_yp = self._labeled_entry(reg, "air_yp (+Y, mm):", "10.0")
        self.entry_air_zm = self._labeled_entry(reg, "air_zm (−Z, mm):", "20.0")
        self.entry_air_zp = self._labeled_entry(reg, "air_zp (+Z, mm):", "20.0")

    def _build_tab_solve(self):
        grid = ctk.CTkFrame(self.tab_solve)
        grid.pack(fill="x", padx=12, pady=12)

        fset = self._section(grid, "Setup em GHz")
        self.entry_f0 = self._labeled_entry(fset, "Frequência do Setup (GHz):", "10.0")

        swp = self._section(grid, "Sweep Linear / Interpolating (GHz)")
        self.entry_fs = self._labeled_entry(swp, "Início fs:", "8.0")
        self.entry_fe = self._labeled_entry(swp, "Fim fe:", "12.0")
        self.entry_fd = self._labeled_entry(swp, "Passo fd:", "0.05")

    # ---- Helpers UI ----
    def _section(self, master, title):
        frame = ctk.CTkFrame(master, corner_radius=14)
        frame.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8, pady=(8,4))
        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack(fill="x", padx=8, pady=(0,8))
        return inner

    def _labeled_entry(self, master, label, default=""):
        row = ctk.CTkFrame(master, fg_color="transparent")
        row.pack(fill="x", pady=4)
        ctk.CTkLabel(row, text=label, width=230, anchor="w").pack(side="left", padx=(0,8))
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
        state = "disabled" if running else "normal"
        self.btn_run.configure(state=state)
        self.btn_save.configure(state=state)
        self.btn_open.configure(state=state)
        self.progress.set(0.35 if running else 0.0)
        self.status.configure(text="Processando..." if running else "Pronto")

    # ---- Ações ----
    def collect_params(self):
        try:
            return {
                "project": self.entry_project.get().strip() or DEFAULT_PROJECT,
                "design":  self.entry_design.get().strip() or DEFAULT_DESIGN,
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
                "fechar_ao_sair": True if self.chk_close.get() else False,
            }
        except ValueError as e:
            raise RuntimeError(f"Parâmetro inválido: {e}")

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
        folder = filedialog.askdirectory(title="Escolha a pasta para salvar")
        if not folder:
            return
        messagebox.showinfo(
            "Salvar .aedt",
            "Após a simulação, use File > Save no HFSS para gravar o projeto na pasta escolhida."
        )

    def on_open_workdir(self):
        try:
            base = os.path.expanduser("~")
            if sys.platform.startswith("win"):
                os.startfile(base)
            elif sys.platform == "darwin":
                os.system(f'open "{base}"')
            else:
                os.system(f'xdg-open "{base}"')
        except Exception as e:
            messagebox.showwarning("Falha ao abrir pasta", str(e))

# Entrypoint
def main():
    app = PatchGUI()
    app.mainloop()

if __name__ == "__main__":
    main()
