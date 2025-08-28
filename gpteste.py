# -*- coding: utf-8 -*-
"""
HFSS Array de Patches — Probe-fed (Terminal) via Stackup3D + GUI
Requisitos: PyAEDT==0.18.1 | HFSS 2024 R2 | customtkinter
"""

from __future__ import annotations
import os, math, shutil, traceback, tempfile, time
from typing import Tuple, List
import customtkinter as ctk

import ansys.aedt.core
from ansys.aedt.core import Hfss
from ansys.aedt.core.modeler.advanced_cad.stackup_3d import Stackup3D

# ===================== Defaults / Constantes =====================
AEDT_VERSION_DEFAULT = "2024.2"
UNITS_LEN = "mm"
UNITS_FREQ = "GHz"
COPPER_T_DEFAULT = 0.035  # mm
# Materiais comuns já presentes na lib do HFSS (pode digitar outro nome na GUI)
COMMON_DK = {
    "FR4_epoxy": 4.4,
    "Duroid (tm)": 2.2,
    "Rogers RO4003C (tm)": 3.55,
    "Rogers RO4350B (tm)": 3.66,
}

# ===================== Utilidades Eletromag =====================
def c_mm_per_GHz() -> float:
    return 299.792458

def _eps_eff_line(eps_r: float, w_h: float) -> float:
    # εeff aproximado de microstrip (usado só p/ λg do λ/4 caso precise)
    return (eps_r + 1)/2 + (eps_r - 1)/2 * (1 + 12/w_h) ** -0.5

def hammerstad_patch_dims(f0_GHz: float, eps_r: float, h_mm: float) -> Tuple[float, float, float]:
    # Dimensões clássicas (Hammerstad/Jensen) para patch retangular em ϵr, h
    c = c_mm_per_GHz()
    W = c/(2.0*f0_GHz)*math.sqrt(2.0/(eps_r+1.0))
    eps_eff = (eps_r + 1.0)/2.0 + (eps_r - 1.0)/2.0 * (1.0/math.sqrt(1.0 + 12.0*h_mm/W))
    dL = 0.412*h_mm*((eps_eff+0.3)*(W/h_mm + 0.264))/((eps_eff-0.258)*(W/h_mm + 0.8))
    L_eff = c/(2.0*f0_GHz*math.sqrt(eps_eff))
    L = L_eff - 2.0*dL
    return W, L, eps_eff

def nx_ny_from_gain(gtar_dbi: float, gelem_dbi: float) -> tuple[int, int, int]:
    # potência ∝ N => Garray ≈ Gelem + 10log10(N)  →  N≈10^((Garray−Gelem)/10)
    need = max(1.0, 10 ** ((gtar_dbi - gelem_dbi) / 10.0))
    n = max(1, math.ceil(need))
    nx = int(round(math.sqrt(n)))
    ny = int(math.ceil(n / nx))
    return nx, ny, nx*ny

def clean_previous(project_path: str):
    # Remove artefatos/locks se existirem
    try:
        if os.path.exists(project_path):
            try: os.remove(project_path)
            except Exception: pass
        if os.path.exists(project_path + ".lock"):
            try: os.remove(project_path + ".lock")
            except Exception: pass
        res_dir = project_path.replace(".aedt", ".aedtresults")
        if os.path.exists(res_dir):
            try: shutil.rmtree(res_dir)
            except Exception: pass
        sem1 = os.path.join(os.path.dirname(project_path), ".PatchArray_HFSS.asol_priv.semaphore")
        sem2 = os.path.join(project_path.replace(".aedt", ".aedtresults"), ".PatchArray_HFSS.asol_priv.semaphore")
        for f in (sem1, sem2):
            if os.path.exists(f):
                try: os.remove(f)
                except Exception: pass
    except Exception:
        pass

# ===================== Builder principal =====================
def build_array_probe_gui(
    fmin_GHz: float,
    fmax_GHz: float,
    diel_name: str,
    h_mm: float,
    t_cu_mm: float,
    use_gain_sizing: bool,
    g_target_dbi: float,
    g_elem_dbi: float,
    nx_in: int,
    ny_in: int,
    pitch_x_mm: float | None,
    pitch_y_mm: float | None,
    feed_rel_x: float,
    region_pad_mm: float,
    aedt_version: str,
    non_graphical: bool,
    solve_after: bool,
    out_dir: str | None = None,
) -> tuple[Hfss, dict]:
    """
    Cria projeto HFSS (Terminal) com array de patches probe-fed (portas individuais),
    a partir dos parâmetros informados na GUI.
    """

    # ======= Frequências e pitch =======
    f0 = 0.5*(fmin_GHz + fmax_GHz)
    lam0 = c_mm_per_GHz() / f0
    px = pitch_x_mm if pitch_x_mm and pitch_x_mm > 0 else 0.5*lam0
    py = pitch_y_mm if pitch_y_mm and pitch_y_mm > 0 else 0.5*lam0

    # ======= Dimensões do patch (Hammerstad) =======
    # Obs.: usamos εr nominal p/ cálculo inicial; refino por otimização é passo seguinte
    epsr_guess = COMMON_DK.get(diel_name, None)
    if epsr_guess is None:
        # Palpite prudente se o material não estiver no dict:
        epsr_guess = 3.0 if "Rogers" in diel_name else 4.4
    Wp, Lp, eps_eff = hammerstad_patch_dims(f0, epsr_guess, h_mm)

    # ======= Tamanho do array =======
    if use_gain_sizing:
        nx, ny, N = nx_ny_from_gain(g_target_dbi, g_elem_dbi)
    else:
        nx = max(1, int(nx_in))
        ny = max(1, int(ny_in))
        N = nx*ny

    # ======= Extensão XY para checar região =======
    sx = (nx - 1)*px + Wp
    sy = (ny - 1)*py + Lp

    # ======= Paths / Projeto =======
    if out_dir is None:
        out_dir = tempfile.mkdtemp(suffix=".ansys")
    os.makedirs(out_dir, exist_ok=True)
    project_path = os.path.join(out_dir, "patch_array_probe.aedt")
    clean_previous(project_path)

    # ======= Lançar HFSS =======
    hfss = Hfss(
        project=project_path,
        design="PatchArray_HFSS",
        solution_type="Terminal",
        new_desktop=True,
        non_graphical=non_graphical,
        version=aedt_version or AEDT_VERSION_DEFAULT,
        remove_lock=True,
    )
    hfss.modeler.model_units = UNITS_LEN

    # ======= Stackup =======
    stack = Stackup3D(hfss)
    ground = stack.add_ground_layer("ground", material="copper", thickness=t_cu_mm, fill_material="air")
    diel = stack.add_dielectric_layer("dielectric", thickness=f"{h_mm}{UNITS_LEN}", material=diel_name)
    signal = stack.add_signal_layer("signal", material="copper", thickness=t_cu_mm, fill_material="air")

    # ======= Helper: criar 1 patch em (ix,iy) =======
    def add_patch_at(ix: int, iy: int, idx: int) -> str:
        name = f"Patch_{ix+1}_{iy+1}"
        # Cria patch nas dimensões analíticas
        p = signal.add_patch(
            patch_length=Lp,
            patch_width=Wp,
            patch_name=name,
            frequency=f0*1e9,  # requer Hz
        )
        # Move p/ posição do grid (centro no (0,0))
        x = (ix - (nx - 1)/2.0) * px
        y = (iy - (ny - 1)/2.0) * py
        try:
            hfss.modeler.move([name], [x, y, 0.0])
        except Exception:
            hfss.modeler.move(name, [x, y, 0.0])

        # Cria porta por probe (Terminal) com offset relativo na largura
        # A API cria o via/terminal conectando top->ground
        p.create_probe_port(ground, rel_x_offset=float(feed_rel_x))

        # Renomeia a última porta para manter ordem (se possível)
        try:
            # Padrão: "port1","port2"...; vamos renomear para P{idx}
            bnds = list(hfss.boundaries.boundaries.keys())
            # último item costuma ser a porta recém-criada
            last = bnds[-1] if bnds else None
            if last and last.lower().startswith("port"):
                newname = f"P{idx}"
                hfss.boundaries.rename_boundary(last, newname)
                return newname
        except Exception:
            pass
        return f"P{idx}"

    # ======= Criar todos os patches/portas =======
    port_names: List[str] = []
    idx = 1
    for ix in range(nx):
        for iy in range(ny):
            pname = add_patch_at(ix, iy, idx)
            port_names.append(pname)
            idx += 1

    # ======= Região + Rad =======
    # Margem suficiente nas 6 faces; o Stackup cuidará de Z automaticamente via espessuras
    pad = [region_pad_mm]*6
    region = hfss.modeler.create_region(pad, is_percentage=False)
    hfss.assign_radiation_boundary_to_objects(region)

    # ======= Setup + Sweep =======
    setup = hfss.create_setup(name="Setup1", setup_type="HFSSDriven", Frequency=f"{f0}{UNITS_FREQ}")
    setup.create_frequency_sweep(
        unit=UNITS_FREQ,
        name="Sweep1",
        start_frequency=max(0.01, fmin_GHz),
        stop_frequency=fmax_GHz,
        sweep_type="Interpolating",
    )

    hfss.save_project()

    # ======= Análise =======
    if solve_after:
        hfss.analyze()

        # Plot rápido (opcional)
        try:
            traces = hfss.get_traces_for_plot()
            if traces:
                report = hfss.post.create_report(traces)
                _ = report.get_solution_data().plot(report.expressions)
        except Exception:
            pass

        hfss.save_project()

    info = {
        "project_path": project_path,
        "f0_GHz": f0,
        "Wp_mm": Wp, "Lp_mm": Lp, "eps_eff_est": eps_eff,
        "nx": nx, "ny": ny, "N": N,
        "pitch_x_mm": px, "pitch_y_mm": py,
        "ports": port_names,
        "region_pad_mm": region_pad_mm,
        "dielectric": diel_name, "h_mm": h_mm, "t_cu_mm": t_cu_mm,
    }
    return hfss, info

# ===================== GUI =====================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HFSS Array de Patches — Probe-fed (Stackup3D)")
        self.geometry("920x680")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.hfss_ref: Hfss | None = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        r = 0
        self._mk_row(r, "Frequência mínima (GHz):", "9.0"); r += 1
        self._mk_row(r, "Frequência máxima (GHz):", "11.0"); r += 1

        # Dielétrico / Espessuras
        self._mk_row(r, "Material dielétrico (catálogo HFSS):", "Duroid (tm)"); r += 1
        self._mk_row(r, "Altura do dielétrico h (mm):", "0.5"); r += 1
        self._mk_row(r, "Espessura do cobre (mm):", f"{COPPER_T_DEFAULT}"); r += 1

        # Dimensionamento Nx×Ny
        self.var_use_gain = ctk.BooleanVar(value=True)
        chk = ctk.CTkCheckBox(self, text="Dimensionar Nx×Ny por ganho alvo", variable=self.var_use_gain, command=self._toggle_gain_mode)
        chk.grid(row=r, column=0, columnspan=2, padx=10, pady=(4, 4), sticky="w"); r += 1
        self._mk_row(r, "Ganho alvo do array (dBi):", "12"); r += 1
        self._mk_row(r, "Ganho de um elemento (dBi):", "7.0"); r += 1
        self._mk_row(r, "Nx (se não usar ganho alvo):", "2"); r += 1
        self._mk_row(r, "Ny (se não usar ganho alvo):", "2"); r += 1

        # Passo / Feed
        self._mk_row(r, "Pitch X (mm) [vazio=λ0/2]:", ""); r += 1
        self._mk_row(r, "Pitch Y (mm) [vazio=λ0/2]:", ""); r += 1
        self._mk_row(r, "Offset relativo do feed em X (0–1):", "0.485"); r += 1

        # Região e opções
        self._mk_row(r, "Margem da região de ar (mm):", "5.0"); r += 1
        self._mk_row(r, "Versão AEDT:", AEDT_VERSION_DEFAULT); r += 1

        self.var_run = ctk.BooleanVar(value=True)
        self.var_ng = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self, text="Rodar simulação após criar", variable=self.var_run).grid(row=r, column=0, padx=10, pady=(4, 4), sticky="w")
        ctk.CTkCheckBox(self, text="Non-graphical (sem UI)", variable=self.var_ng).grid(row=r, column=1, padx=10, pady=(4, 4), sticky="w")
        r += 1

        self.btn = ctk.CTkButton(self, text="Criar Array no HFSS", command=self.on_create)
        self.btn.grid(row=r, column=0, padx=10, pady=(6, 10), sticky="w")
        r += 1

        self.txt = ctk.CTkTextbox(self, width=880, height=320)
        self.txt.grid(row=r, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(r, weight=1)

        self._log("[Info] Interface baseada no exemplo oficial (Stackup3D + probe).")
        self._log("[Info] Cada elemento recebe sua própria porta Terminal via create_probe_port().")

        # Ajusta estado inicial de Nx/Ny
        self._toggle_gain_mode()

    def _mk_row(self, r, label, default):
        ctk.CTkLabel(self, text=label).grid(row=r, column=0, padx=10, pady=4, sticky="e")
        e = ctk.CTkEntry(self)
        e.insert(0, default)
        e.grid(row=r, column=1, padx=10, pady=4, sticky="ew")
        setattr(self, f"e{r}", e)

    def _toggle_gain_mode(self):
        use_gain = self.var_use_gain.get()
        # Campos Nx/Ny na ordem: (depois do campo "Ganho de um elemento")
        # De acordo com criação acima: indices (0..)
        # fmin=0,fmax=1,diel=2,h=3,tcu=4,chk=5,gtar=6,gelem=7,Nx=8,Ny=9,px=10,py=11,offset=12,region=13,ver=14
        self.e8.configure(state="disabled" if use_gain else "normal")
        self.e9.configure(state="disabled" if use_gain else "normal")

    def _log(self, s: str):
        self.txt.insert("end", s + "\n")
        self.txt.see("end")

    def on_create(self):
        try:
            # Mapear indices dos inputs
            fmin = float(self.e0.get())
            fmax = float(self.e1.get())
            if fmax <= fmin:
                raise ValueError("fmax deve ser maior que fmin.")

            diel = self.e2.get().strip()
            h = float(self.e3.get())
            tcu = float(self.e4.get())

            use_gain = self.var_use_gain.get()
            gtar = float(self.e6.get())
            gelem = float(self.e7.get())
            nx = int(self.e8.get()) if self.e8.get().strip() else 1
            ny = int(self.e9.get()) if self.e9.get().strip() else 1

            px = float(self.e10.get()) if self.e10.get().strip() else None
            py = float(self.e11.get()) if self.e11.get().strip() else None
            feed_rel_x = float(self.e12.get())
            region_pad = float(self.e13.get())
            aedt_ver = self.e14.get().strip() or AEDT_VERSION_DEFAULT

            if not (0.0 <= feed_rel_x <= 1.0):
                raise ValueError("Offset relativo do feed deve estar entre 0 e 1.")

            run_after = self.var_run.get()
            ng = self.var_ng.get()

            f0 = 0.5*(fmin+fmax)
            epsr_guess = COMMON_DK.get(diel, 4.4)
            Wp, Lp, ee = hammerstad_patch_dims(f0, epsr_guess, h)
            self._log(f"[Analítico] f0={f0:.3f} GHz | W≈{Wp:.2f} mm, L≈{Lp:.2f} mm, εeff(est)≈{ee:.4f}")

            out_dir = os.path.dirname(os.path.abspath(__file__))  # salva no diretório do script

            hfss, info = build_array_probe_gui(
                fmin_GHz=fmin, fmax_GHz=fmax,
                diel_name=diel, h_mm=h, t_cu_mm=tcu,
                use_gain_sizing=use_gain, g_target_dbi=gtar, g_elem_dbi=gelem,
                nx_in=nx, ny_in=ny,
                pitch_x_mm=px, pitch_y_mm=py,
                feed_rel_x=feed_rel_x,
                region_pad_mm=region_pad,
                aedt_version=aedt_ver,
                non_graphical=ng,
                solve_after=run_after,
                out_dir=out_dir,
            )
            self.hfss_ref = hfss

            self._log(f"[Projeto] {info['project_path']}")
            self._log(f"[Array] {info['nx']}×{info['ny']} = {info['N']}  | pitch=({info['pitch_x_mm']:.2f},{info['pitch_y_mm']:.2f}) mm")
            self._log(f"[Patch] W≈{info['Wp_mm']:.2f} mm, L≈{info['Lp_mm']:.2f} mm | h={info['h_mm']:.3f} mm | diel='{info['dielectric']}'")
            self._log(f"[Ports] {', '.join(info['ports'])}")
            self._log(f"[Boundary] Radiation em região com margem {info['region_pad_mm']:.2f} mm")
            self._log("[Status] OK — modelo criado. Se 'Rodar simulação' estava marcado, o setup já foi executado.")

        except Exception as e:
            self._log("Erro: " + str(e))
            self._log(traceback.format_exc())

    def on_close(self):
        try:
            if self.hfss_ref is not None:
                # Tenta salvar e liberar
                try:
                    self.hfss_ref.save_project()
                except Exception:
                    pass
                self.hfss_ref.release_desktop(close_projects=True, close_desktop=True)
                self._log("[AEDT] Instância liberada.")
        except Exception as e:
            self._log(f"[AEDT] Falha ao liberar: {e}")
        finally:
            self.destroy()

# ===================== Main =====================
if __name__ == "__main__":
    app = App()
    app.mainloop()
