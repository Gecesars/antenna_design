# -*- coding: utf-8 -*-
"""
HFSS Patch Array — λ/4 Transformer + Port Sheet XZ, Setup explícito
PyAEDT==0.18.1 | HFSS 2024 R2
"""

from __future__ import annotations
import os, math, shutil, traceback
from typing import Tuple, List, Dict, Any
import tempfile

import customtkinter as ctk
from ansys.aedt.core import Hfss

# ===================== Configurações Gerais =====================
AEDT_VERSION = "2024.2"
UNITS = "mm"
COPPER_T = 0.035
PATCH_GAIN_DBI = 6.5
Z0_PORT = 50.0
ZPATCH_EDGE = 200.0
PAD_MIN = 1.0

SETUP_NAME = "OpSetup"
SWEEP_NAME = "OpFast"

# ===================== Utilidades Eletromag =====================
def c_mm_per_GHz() -> float:
    return 299.792458

def hammerstad_patch_dims(f0_GHz: float, eps_r: float, h_mm: float) -> Tuple[float, float, float]:
    c = c_mm_per_GHz()
    W = c / (2.0 * f0_GHz) * math.sqrt(2.0 / (eps_r + 1.0))
    eps_eff = (eps_r + 1.0) / 2.0 + (eps_r - 1.0) / 2.0 * (1.0 / math.sqrt(1.0 + 12.0 * h_mm / W))
    dL = 0.412 * h_mm * ((eps_eff + 0.3) * (W/h_mm + 0.264)) / ((eps_eff - 0.258) * (W/h_mm + 0.8))
    L_eff = c / (2.0 * f0_GHz * math.sqrt(eps_eff))
    L = L_eff - 2.0 * dL
    return W, L, eps_eff

def _eps_eff_line(eps_r: float, w_h: float) -> float:
    return (eps_r + 1)/2 + (eps_r - 1)/2 * (1 + 12/w_h) ** -0.5

def z0_from_w_h(eps_r: float, w_h: float) -> float:
    ee = _eps_eff_line(eps_r, w_h)
    if w_h <= 1:
        return (60 / math.sqrt(ee)) * math.log(8.0 / w_h + 0.25 * w_h)
    else:
        return (120 * math.pi) / (math.sqrt(ee) * (w_h + 1.393 + 0.667 * math.log(w_h + 1.444)))

def w_for_z0(eps_r: float, h_mm: float, z0_ohm: float) -> float:
    lo, hi = 0.05, 20.0
    tolerance = 0.001
    max_iter = 50
    
    for _ in range(max_iter):
        mid = (lo + hi) / 2.0
        z0_calc = z0_from_w_h(eps_r, mid)
        if abs(z0_calc - z0_ohm) < tolerance:
            break
        if z0_calc > z0_ohm:
            lo = mid
        else:
            hi = mid
    return mid * h_mm

def quarter_wave_len_mm(f0_GHz: float, eps_r: float, w_mm: float, h_mm: float) -> float:
    w_h = max(w_mm / h_mm, 1e-6)
    ee = _eps_eff_line(eps_r, w_h)
    lambda_g = c_mm_per_GHz() / (f0_GHz * math.sqrt(ee))
    return lambda_g / 4.0

# ===================== Arquivos / Projeto ======================
def clean_previous(project_path: str):
    if os.path.exists(project_path):
        try: 
            os.remove(project_path)
        except Exception: 
            pass
            
    if os.path.exists(project_path + ".lock"):
        try: 
            os.remove(project_path + ".lock")
        except Exception: 
            pass
            
    res_dir = project_path.replace(".aedt", ".aedtresults")
    if os.path.exists(res_dir):
        try: 
            shutil.rmtree(res_dir)
        except Exception: 
            pass

    semaphore_files = [
        os.path.join(os.path.dirname(project_path), ".PatchArray_HFSS.asol_priv.semaphore"),
        os.path.join(res_dir, ".PatchArray_HFSS.asol_priv.semaphore")
    ]
    
    for sem_file in semaphore_files:
        if os.path.exists(sem_file):
            try:
                os.remove(sem_file)
            except Exception:
                pass

# ===================== Construção do Modelo =====================
def create_param(hfss: Hfss, name: str, expr: str) -> str:
    hfss[name] = expr
    return name

def build_array_project(
    fmin_GHz: float,
    fmax_GHz: float,
    g_target_dbi: float,
    eps_r: float,
    h_mm: float,
    out_dir: str,
    solve_after: bool = True,
) -> tuple[Hfss, dict]:
    os.makedirs(out_dir, exist_ok=True)
    project_path = os.path.join(out_dir, "patch_array_hfss_modal.aedt")
    clean_previous(project_path)

    f0 = 0.5 * (fmin_GHz + fmax_GHz)
    delta = 0.5 * (fmax_GHz - fmin_GHz) * 1.3
    fstart = max(0.01, f0 - delta)
    fstop = f0 + delta

    Wp, Lp, eps_eff_patch = hammerstad_patch_dims(f0, eps_r, h_mm)

    need = max(1.0, 10 ** ((g_target_dbi - PATCH_GAIN_DBI) / 10.0))
    n = max(1, math.ceil(need))
    nx = int(round(math.sqrt(n)))
    ny = int(math.ceil(n / nx))
    N = nx * ny

    Zt = math.sqrt(Z0_PORT * ZPATCH_EDGE)
    Wt = w_for_z0(eps_r, h_mm, Zt)
    Lq = quarter_wave_len_mm(f0, eps_r, Wt, h_mm)
    Pad = max(PAD_MIN, 0.6 * Wt)

    lam0 = c_mm_per_GHz() / f0
    pitch = 0.5 * lam0
    margin_x = max(0.25 * pitch, 10.0)
    margin_y = max(0.25 * pitch, 10.0, Lq + Pad + 5.0)
    sx = nx * pitch + 2 * margin_x
    sy = ny * pitch + 2 * margin_y

    hfss = Hfss(
        project=project_path,
        design="PatchArray_HFSS",
        solution_type="Modal",
        new_desktop=True,
        non_graphical=False,
        version=AEDT_VERSION,
        remove_lock=True
    )
    hfss.modeler.model_units = UNITS

    create_param(hfss, "f0", f"{f0:.6f}GHz")
    create_param(hfss, "epsr", f"{eps_r}")
    create_param(hfss, "h", f"{h_mm:.6f}mm")
    create_param(hfss, "Wp", f"{Wp:.6f}mm")
    create_param(hfss, "Lp", f"{Lp:.6f}mm")
    create_param(hfss, "Wt", f"{Wt:.6f}mm")
    create_param(hfss, "Lq", f"{Lq:.6f}mm")
    create_param(hfss, "Pad", f"{Pad:.6f}mm")
    create_param(hfss, "pitch", f"{pitch:.6f}mm")
    create_param(hfss, "sx", f"{sx:.6f}mm")
    create_param(hfss, "sy", f"{sy:.6f}mm")
    create_param(hfss, "tCu", f"{COPPER_T:.6f}mm")

    gnd = hfss.modeler.create_box(
        origin=[f"-sx/2", f"-sy/2", "0"],
        sizes=["sx", "sy", "tCu"],
        name="GND", 
        material="pec"
    )
    
    sub = hfss.modeler.create_box(
        origin=[f"-sx/2", f"-sy/2", "tCu"],
        sizes=["sx", "sy", "h"],
        name="SUB", 
        material="FR4_epoxy"
    )
    
    z_top = "tCu+h"

    port_names: List[str] = []
    x0 = -(nx - 1) * pitch / 2.0
    y0 = -(ny - 1) * pitch / 2.0

    for ix in range(nx):
        for iy in range(ny):
            cx = x0 + ix * pitch
            cy = y0 + iy * pitch

            cx_s = f"{cx:.6f}mm"
            cy_s = f"{cy:.6f}mm"

            patch = hfss.modeler.create_box(
                origin=[f"{cx_s}-Wp/2", f"{cy_s}-Lp/2", z_top],
                sizes=["Wp", "Lp", "tCu"],
                name=f"PATCH_{ix+1}_{iy+1}", 
                material="pec"
            )

            patch_y_min = f"{cy_s}-Lp/2"
            line_y0 = f"({patch_y_min})-Lq"
            feed_y0 = f"({line_y0})-Pad"

            line = hfss.modeler.create_box(
                origin=[f"{cx_s}-Wt/2", line_y0, z_top],
                sizes=["Wt", "Lq", "tCu"],
                name=f"TPLINE_{ix+1}_{iy+1}", 
                material="pec"
            )
            
            pad = hfss.modeler.create_box(
                origin=[f"{cx_s}-Wt/2", feed_y0, z_top],
                sizes=["Wt", "Pad", "tCu"],
                name=f"FEEDPAD_{ix+1}_{iy+1}", 
                material="pec"
            )

            # Criação correta do sheet para a porta
            sheet_name = f"PORTSHEET_{ix+1}_{iy+1}"
            port_sheet = hfss.modeler.create_rectangle(
                cs_plane="XZ",
                position=[f"{cx_s}-Wt/2", feed_y0, "0"],
                dimension_list=["Wt", "h"],
                name=sheet_name
            )
            
            pname = f"P{len(port_names)+1}"
            
            # Cria a porta lumped usando faces específicas
            faces = hfss.modeler.get_object_faces(port_sheet.name)
            if faces:
                port_face_id = faces[0]
                port = hfss.lumped_port(
                    assignment=port_face_id,
                    reference=gnd,
                    impedance=Z0_PORT, 
                    renormalize=True, 
                    name=pname
                )
                port_names.append(pname)

    hfss.create_open_region(frequency=f"{f0}GHz")

    for setup_name in list(hfss.setups.keys()):
        if setup_name != SETUP_NAME:
            hfss.setups[setup_name].delete()

    setup = hfss.create_setup(setupname=SETUP_NAME)
    setup.props["Frequency"] = f"{f0}GHz"
    setup.props["MaximumPasses"] = 6
    setup.props["MaximumDeltaS"] = 0.02
    setup.update()
    
    sweep = setup.create_frequency_sweep(
        sweepname=SWEEP_NAME,
        unit="GHz",
        freqstart=fstart,
        freqstop=fstop,
        sweep_type="Interpolating",
        num_of_freq_points=101
    )

    hfss.save_project()

    if solve_after:
        hfss.analyze_setup(SETUP_NAME)

        try:
            report = hfss.post.reports_by_category.standard("dB(S(1,1))")
            report.props["Primary Sweep"] = "Freq"
            report.props["Secondary Sweep"] = ""
            report.create()
        except Exception as e:
            print(f"Erro ao criar relatório: {e}")

        hfss.save_project()

    info = {
        "project_path": project_path,
        "f0_GHz": f0,
        "Wp_mm": Wp, "Lp_mm": Lp, "eps_eff_patch": eps_eff_patch,
        "Zt_ohm": Zt, "Wt_mm": Wt, "Lq_mm": Lq, "Pad_mm": Pad,
        "nx": nx, "ny": ny, "N": N,
        "ports": port_names,
        "setup": SETUP_NAME, "sweep": SWEEP_NAME,
        "fstart_GHz": fstart, "fstop_GHz": fstop
    }
    return hfss, info

# ===================== GUI (customtkinter) =====================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HFSS Patch Array — λ/4 (PyAEDT)")
        self.geometry("820x600")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.hfss_ref: Hfss | None = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._mk_row(0, "Frequência mínima (GHz):", "2.3")
        self._mk_row(1, "Frequência máxima (GHz):", "2.5")
        self._mk_row(2, "Ganho alvo do array (dBi):", "12")
        self._mk_row(3, "εr (FR4≈4.4):", "4.4")
        self._mk_row(4, "Altura do substrato h (mm):", "1.57")

        self.chk_run = ctk.CTkCheckBox(self, text="Rodar simulação após criar")
        self.chk_run.grid(row=5, column=1, padx=10, pady=(6, 6), sticky="w")
        self.chk_run.select()

        self.btn = ctk.CTkButton(self, text="Criar Array no HFSS", command=self.on_create)
        self.btn.grid(row=6, column=1, padx=10, pady=(0, 8), sticky="w")

        self.txt = ctk.CTkTextbox(self, width=780, height=360)
        self.txt.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(7, weight=1)
        self._log("[Dica] Porta lumped criada usando sheets XZ.")
        self._log("[Dica] Setup executado explicitamente: OpSetup : OpFast.")

    def _mk_row(self, r, label, default):
        ctk.CTkLabel(self, text=label).grid(row=r, column=0, padx=10, pady=6, sticky="e")
        e = ctk.CTkEntry(self)
        e.insert(0, default)
        e.grid(row=r, column=1, padx=10, pady=6, sticky="ew")
        setattr(self, f"e{r}", e)

    def _log(self, s: str):
        self.txt.insert("end", s + "\n")
        self.txt.see("end")

    def on_create(self):
        try:
            fmin = float(self.e0.get())
            fmax = float(self.e1.get())
            if fmax <= fmin: 
                raise ValueError("fmax deve ser maior que fmin.")
                
            gtar = float(self.e2.get())
            epsr = float(self.e3.get())
            h = float(self.e4.get())
            run_after = self.chk_run.get()

            f0 = 0.5*(fmin+fmax)
            Wp, Lp, ee = hammerstad_patch_dims(f0, epsr, h)
            self._log(f"[Analítico] f0={f0:.3f} GHz | W≈{Wp:.2f} mm, L≈{Lp:.2f} mm, εeff≈{ee:.4f}")

            script_dir = os.path.dirname(os.path.abspath(__file__))
            out_dir = script_dir

            hfss, info = build_array_project(
                fmin_GHz=fmin, fmax_GHz=fmax, g_target_dbi=gtar,
                eps_r=epsr, h_mm=h, out_dir=out_dir, solve_after=run_after
            )
            self.hfss_ref = hfss

            self._log(f"[Projeto] {info['project_path']}")
            self._log(f"[Linha λ/4] Zt≈{info['Zt_ohm']:.1f} Ω | Wt≈{info['Wt_mm']:.2f} mm | Lq≈{info['Lq_mm']:.2f} mm | Pad≈{info['Pad_mm']:.2f} mm")
            self._log(f"[Sweep] {info['setup']} : {info['sweep']}  ({info['fstart_GHz']:.3f}–{info['fstop_GHz']:.3f} GHz)")
            self._log(f"[Ports] {', '.join(info['ports'])}")
            self._log(f"[Array] {info['nx']}×{info['ny']} = {info['N']} elementos")
            
        except Exception as e:
            self._log("Erro: " + str(e))
            self._log(traceback.format_exc())

    def on_close(self):
        try:
            if self.hfss_ref is not None:
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