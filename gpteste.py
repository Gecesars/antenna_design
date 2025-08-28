# -*- coding: utf-8 -*-
"""
HFSS Patch Array (Stackup3D, probe-fed) — GUI
PyAEDT==0.18.1 | HFSS 2024 R2
"""

from __future__ import annotations
import os, math, shutil, traceback, tempfile
from typing import Tuple, List, Dict, Any

import customtkinter as ctk
from ansys.aedt.core import Hfss
from ansys.aedt.core.modeler.advanced_cad.stackup_3d import Stackup3D

# ===================== Constantes =====================
AEDT_VERSION = "2024.2"
UNITS = "mm"
COPPER_T = 0.035      # mm
Z0_TARGET = 50.0      # Ohm
R_EDGE = 200.0        # Ohm (aprox. resistência na borda do patch)
MARGIN_MM = 15.0      # região de ar ao redor
PROBE_DIAM_MM = 1.2   # diâmetro do pino

SETUP_NAME = "Setup_Array"
SWEEP_NAME = "Sweep1"

# ===================== EM utils =====================
def c_mm_per_GHz() -> float:
    return 299.792458  # mm/GHz

def hammerstad_patch_dims(f0_GHz: float, eps_r: float, h_mm: float) -> Tuple[float, float, float]:
    """W, L, eps_eff (Hammerstad)."""
    c = c_mm_per_GHz()
    W = c / (2.0 * f0_GHz) * math.sqrt(2.0 / (eps_r + 1.0))
    eps_eff = (eps_r + 1.0) / 2.0 + (eps_r - 1.0) / 2.0 * (1.0 / math.sqrt(1.0 + 12.0 * h_mm / W))
    dL = 0.412 * h_mm * ((eps_eff + 0.3) * (W/h_mm + 0.264)) / ((eps_eff - 0.258) * (W/h_mm + 0.8))
    L_eff = c / (2.0 * f0_GHz * math.sqrt(eps_eff))
    L = L_eff - 2.0 * dL
    return W, L, eps_eff

def rel_x_offset_for_impedance(Zedge: float, Z0: float) -> float:
    """
    Para alimentação ao longo do eixo do comprimento:
    Rin(y0) = Zedge * cos^2(pi*y0/L), com y0=0 na borda.
    Converter para offset relativo do centro (rel_x_offset):
      rel = 1 - 2*y0/L.  (0=centro, 1=borda)
    """
    if Z0 >= Zedge:
        return 1.0  # empurra para a borda (limite)
    y0_over_L = (1.0 / math.pi) * math.acos(math.sqrt(Z0 / Zedge))
    rel = 1.0 - 2.0 * y0_over_L
    return max(0.0, min(1.0, rel))

# ===================== Limpeza =====================
def clean_previous(project_path: str):
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
    # semáforos ocasionais
    sem = os.path.join(os.path.dirname(project_path), ".PatchArray_HFSS.asol_priv.semaphore")
    if os.path.exists(sem):
        try: os.remove(sem)
        except Exception: pass

# ===================== Construção =====================
def build_array_project(
    fmin_GHz: float,
    fmax_GHz: float,
    eps_r: float,
    h_mm: float,
    nx: int,
    ny: int,
    pitch_override_mm: float | None,
    run_after: bool
) -> tuple[Hfss, dict]:

    f0 = 0.5*(fmin_GHz+fmax_GHz)
    lam0 = c_mm_per_GHz() / f0
    pitch = pitch_override_mm if pitch_override_mm and pitch_override_mm > 0 else 0.5 * lam0

    Wp, Lp, ee = hammerstad_patch_dims(f0, eps_r, h_mm)
    feed_rel = rel_x_offset_for_impedance(R_EDGE, Z0_TARGET)

    # domínio e setup
    out_dir = os.path.dirname(os.path.abspath(__file__))
    project_path = os.path.join(out_dir, "patch_array_stackup.aedt")
    clean_previous(project_path)

    hfss = Hfss(
        project=project_path,
        solution_type="Terminal",
        design="PatchArray_HFSS_Stackup",
        new_desktop=True,
        non_graphical=False,
        version=AEDT_VERSION,
        remove_lock=True
    )
    hfss.modeler.model_units = UNITS

    # Stackup
    stk = Stackup3D(hfss)
    gnd = stk.add_ground_layer("ground", material="copper", thickness=COPPER_T, fill_material="air")
    diel = stk.add_dielectric_layer("dielectric", thickness=f"{h_mm}", material="FR4_epoxy")
    top  = stk.add_signal_layer("signal", material="copper", thickness=COPPER_T, fill_material="air")

    # Geração do array
    # origem no centro do array
    x0 = -(nx - 1) * pitch / 2.0
    y0 = -(ny - 1) * pitch / 2.0

    patches = []
    port_names = []

    for ix in range(nx):
        for iy in range(ny):
            cx = x0 + ix * pitch
            cy = y0 + iy * pitch
            pname = f"P_{ix+1}_{iy+1}"
            patch_name = f"Patch_{ix+1}_{iy+1}"

            # >>> PyAEDT 0.18.1: frequency é o 1º argumento (em Hz!)
            patch = top.add_patch(
                f0 * 1e9,
                patch_width=Wp,
                patch_length=Lp,
                patch_position_x=cx,
                patch_position_y=cy,
                patch_name=patch_name,
                axis="X"
            )
            if not patch:  # add_patch falhou
                raise RuntimeError(f"Falha ao criar {patch_name}. Verifique parâmetros/posição.")

            # Probe individual com nome único
            ok = patch.create_probe_port(
                gnd,
                rel_x_offset=feed_rel,
                rel_y_offset=0.0,
                r=PROBE_DIAM_MM,   # diâmetro do pino (mm) conforme doc do 0.18.1
                name=pname
            )
            if not ok:
                raise RuntimeError(f"Falha ao criar porta (probe) para {patch_name}.")

            patches.append(patch)
            port_names.append(pname + "_T1")  # porta terminal será <name>_T1

    # Região e radiação
    # tamanho total do array:
    size_x = (nx - 1) * pitch + Wp + 2*MARGIN_MM
    size_y = (ny - 1) * pitch + Lp + 2*MARGIN_MM
    # centraliza região
    region = hfss.modeler.create_region(
        [MARGIN_MM, MARGIN_MM, MARGIN_MM, MARGIN_MM, MARGIN_MM, MARGIN_MM],
        is_percentage=False
    )
    hfss.assign_radiation_boundary_to_objects(region)

    # Setup e sweep
    setup = hfss.create_setup(name=SETUP_NAME, setup_type="HFSSDriven", Frequency=f"{f0}GHz")
    setup.create_frequency_sweep(
        unit="GHz",
        name=SWEEP_NAME,
        start_frequency=max(0.01, fmin_GHz),
        stop_frequency=fmax_GHz,
        sweep_type="Interpolating",
    )

    hfss.save_project()

    if run_after:
        hfss.analyze(setups=[SETUP_NAME])

    info = dict(
        project_path=project_path,
        f0_GHz=f0, Wp_mm=Wp, Lp_mm=Lp, eps_eff=ee,
        nx=nx, ny=ny, pitch_mm=pitch,
        ports=port_names, setup=SETUP_NAME, sweep=SWEEP_NAME
    )
    return hfss, info

# ===================== GUI =====================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HFSS Patch Array — Stackup3D (Probe-fed)")
        self.geometry("860x560")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.hfss_ref: Hfss | None = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._mk_row(0, "Frequência mínima (GHz):", "2.3")
        self._mk_row(1, "Frequência máxima (GHz):", "2.5")
        self._mk_row(2, "εr (ex.: FR4≈4.4):", "4.4")
        self._mk_row(3, "Altura do substrato h (mm):", "1.57")
        self._mk_row(4, "Nx (elementos):", "2")
        self._mk_row(5, "Ny (elementos):", "2")
        self._mk_row(6, "Pitch do array (mm) [vazio = λ0/2]:", "")

        self.chk_run = ctk.CTkCheckBox(self, text="Rodar simulação após criar")
        self.chk_run.grid(row=7, column=1, padx=10, pady=(6, 6), sticky="w")
        self.chk_run.select()

        self.btn = ctk.CTkButton(self, text="Criar Array no HFSS", command=self.on_create)
        self.btn.grid(row=8, column=1, padx=10, pady=(0, 8), sticky="w")

        self.txt = ctk.CTkTextbox(self, width=820, height=320)
        self.txt.grid(row=9, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(9, weight=1)

        self._log("[Info] Stackup3D (Terminal) + probe por elemento. 'frequency' passado como 1º arg no add_patch.")

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
            fmin = float(self.e0.get()); fmax = float(self.e1.get())
            if fmax <= fmin: raise ValueError("fmax deve ser maior que fmin.")
            epsr = float(self.e2.get()); h = float(self.e3.get())
            nx = int(self.e4.get()); ny = int(self.e5.get())
            pitch_str = self.e6.get().strip()
            pitch = float(pitch_str) if pitch_str else None
            run_after = self.chk_run.get()

            f0 = 0.5*(fmin+fmax)
            Wp, Lp, ee = hammerstad_patch_dims(f0, epsr, h)
            feed_rel = rel_x_offset_for_impedance(R_EDGE, Z0_TARGET)

            self._log(f"[Analítico] f0={f0:.3f} GHz | W≈{Wp:.2f} mm, L≈{Lp:.2f} mm, εeff≈{ee:.4f}")
            self._log(f"[Probe] Zedge≈{R_EDGE:.1f} Ω → alvo {Z0_TARGET:.1f} Ω → rel_x_offset≈{feed_rel:.3f} (0=centro, 1=borda)")

            hfss, info = build_array_project(
                fmin_GHz=fmin, fmax_GHz=fmax,
                eps_r=epsr, h_mm=h,
                nx=nx, ny=ny,
                pitch_override_mm=pitch,
                run_after=run_after
            )
            self.hfss_ref = hfss

            self._log(f"[Projeto] {info['project_path']}")
            self._log(f"[Array] {nx}×{ny} | pitch={info['pitch_mm']:.2f} mm")
            self._log(f"[Ports] {', '.join(info['ports'])}")
            self._log(f"[Sweep] {info['setup']} : {info['sweep']}  ({fmin:.3f}–{fmax:.3f} GHz)")

        except Exception as e:
            self._log("ERRO: " + str(e))
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
