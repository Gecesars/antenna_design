# -*- coding: utf-8 -*-
"""
HFSS Patch Array – Automação (PyAEDT + GUI)
-------------------------------------------
• HFSS 3D (Driven Modal), Auto Open Region.
• Setup em f0 e sweep "Interpolating/Fast" com 101 pontos.
• Arranjo de patches para ganho-alvo (aprox. Gelem + 10*log10(N)).
• Uma Lumped Port por elemento (via hfss.lumped_port).
• Variáveis Amp_Pi (W) e Phi_Pi (deg) por porta; aplicadas via edit_sources
  usando o formato aceito: {"P1" ou "P1:1": ("Amp_P1", "Phi_P1"), ...}
• GUI (customtkinter) com labels e logs.

Requisitos:
    pip install ansys-aedt-core==0.18.1 customtkinter
    HFSS 2024 R2 instalado/licenciado
"""

from __future__ import annotations
import os
import math
import shutil
import traceback
import customtkinter as ctk
from typing import Tuple, List

from ansys.aedt.core import Hfss

# =========================
# Parâmetros padrão
# =========================
AEDT_VERSION = "2024.2"   # HFSS 2024 R2
UNITS = "mm"
COPPER_T_MM = 0.035
PATCH_GAIN_DBI = 6.5      # ganho típico de patch isolado (aprox.)

# =========================
# Utilitários RF
# =========================
def c_mm_per_GHz() -> float:
    return 299.792458

def hammerstad_patch_dims(f0_GHz: float, eps_r: float, h_mm: float) -> Tuple[float, float, float]:
    """
    Dimensões aproximadas do patch retangular (modo TM10).
    Retorna (W_mm, L_mm, eps_eff)
    """
    c = c_mm_per_GHz()
    W = c / (2.0 * f0_GHz) * math.sqrt(2.0 / (eps_r + 1.0))
    eps_eff = (eps_r + 1.0) / 2.0 + (eps_r - 1.0) / 2.0 * (1.0 / math.sqrt(1 + 12.0 * h_mm / W))
    dL = 0.412 * h_mm * ((eps_eff + 0.3) * (W/h_mm + 0.264)) / ((eps_eff - 0.258) * (W/h_mm + 0.8))
    L_eff = c / (2.0 * f0_GHz * math.sqrt(eps_eff))
    L = L_eff - 2.0 * dL
    return W, L, eps_eff

def estimate_array_layout(g_target_dbi: float, g_elem_dbi: float = PATCH_GAIN_DBI) -> Tuple[int,int,int]:
    """
    N ≈ 10^((Gtarget - Gelem)/10). Retorna grade quase quadrada (nx,ny,n_real).
    """
    need = max(1.0, 10.0 ** ((g_target_dbi - g_elem_dbi) / 10.0))
    n = max(1, math.ceil(need))
    nx = int(round(math.sqrt(n)))
    ny = int(math.ceil(n / nx))
    return nx, ny, nx*ny

def mm(x: float) -> str:
    return f"{x:.6f}mm"

def ghz(x: float) -> str:
    return f"{x:.6f}GHz"

# =========================
# FS helpers
# =========================
def clean_previous(project_path: str):
    if os.path.exists(project_path):
        try: os.remove(project_path)
        except Exception: pass
    if os.path.exists(project_path + ".lock"):
        try: os.remove(project_path + ".lock")
        except Exception: pass
    res = project_path.replace(".aedt", ".aedtresults")
    if os.path.exists(res):
        try: shutil.rmtree(res)
        except Exception: pass

def ensure_amp_phase_vars(hfss: Hfss, port_names: List[str]):
    """Cria Amp_Pi (W) e Phi_Pi (deg) por porta. 1 W total dividido por N."""
    n = max(1, len(port_names))
    p_each = 1.0 / n
    for i in range(1, n+1):
        hfss[f"Amp_P{i}"] = f"{p_each:.6f}W"
        hfss[f"Phi_P{i}"] = "0deg"

def edit_sources_robusto(hfss: Hfss, port_names: List[str]) -> bool:
    """
    Aplica pesos pós-solve. Tenta chaves 'P1' e 'P1:1'.
    Formato aceito: dict[name] = (magnitude_str, phase_str)
    """
    # 1) tenta 'P1'
    try:
        d = {f"P{i}": (f"Amp_P{i}", f"Phi_P{i}") for i in range(1, len(port_names)+1)}
        hfss.edit_sources(d)
        return True
    except Exception:
        pass
    # 2) tenta 'P1:1'
    try:
        d = {f"P{i}:1": (f"Amp_P{i}", f"Phi_P{i}") for i in range(1, len(port_names)+1)}
        hfss.edit_sources(d)
        return True
    except Exception:
        return False

# =========================
# Núcleo: criação do array
# =========================
def create_patch_array_hfss(
    fmin_GHz: float,
    fmax_GHz: float,
    g_target_dbi: float,
    eps_r: float,
    h_mm: float,
    solve_after: bool,
    out_dir: str,
) -> Tuple[Hfss, dict]:
    """
    Constrói projeto HFSS (Driven Modal) com arranjo de patches e retorna (hfss, info).
    """
    os.makedirs(out_dir, exist_ok=True)
    project_path = os.path.join(out_dir, "patch_array_hfss_modal.aedt")
    clean_previous(project_path)

    f0 = 0.5 * (fmin_GHz + fmax_GHz)
    W, L, eps_eff = hammerstad_patch_dims(f0, eps_r, h_mm)
    nx, ny, n_real = estimate_array_layout(g_target_dbi)

    # Passo ~0.5 λ0
    lam0_mm = c_mm_per_GHz() / f0
    pitch = 0.5 * lam0_mm
    sx = nx * pitch + 2*max(0.25*pitch, 10.0)
    sy = ny * pitch + 2*max(0.25*pitch, 10.0)

    # === HFSS ===
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

    # Ground (sólido PEC)
    gnd = hfss.modeler.create_box(
        [-sx/2, -sy/2, 0.0],
        [sx, sy, COPPER_T_MM],
        name="GND",
        matname="pec"
    )

    # Substrato FR4
    sub = hfss.modeler.create_box(
        [-sx/2, -sy/2, COPPER_T_MM],
        [sx, sy, h_mm],
        name="SUB",
        matname="FR4_epoxy"
    )

    z_top = COPPER_T_MM + h_mm  # topo do dielétrico
    port_names: List[str] = []

    # Cria NX x NY patches como sólidos finos (PEC) + lumped_port ao GND
    x0 = -(nx - 1) * pitch / 2.0
    y0 = -(ny - 1) * pitch / 2.0
    for ix in range(nx):
        for iy in range(ny):
            cx = x0 + ix * pitch
            cy = y0 + iy * pitch
            p = hfss.modeler.create_box(
                [cx - W/2.0, cy - L/2.0, z_top],
                [W, L, COPPER_T_MM],
                name=f"PATCH_{ix+1}_{iy+1}",
                matname="pec"
            )
            pname = f"P{len(port_names)+1}"
            # --- Porta Lumped robusta (API estável na 0.18.1) ---
            hfss.lumped_port(
                assignment=p,
                reference=gnd,
                create_port_sheet=True,   # cria folha de porta automaticamente
                port_on_plane=True,       # integração normal ao plano
                impedance=50,
                renormalize=True,
                name=pname
            )
            port_names.append(pname)

    # Auto Open Region (encapsula criação do boundary Radiation)
    hfss.create_open_region(frequency=ghz(f0))

    # Setup + Sweep Interpolating (101 pontos)
    setup = hfss.create_setup(name="Setup1", Frequency=ghz(f0))
    setup.create_frequency_sweep(
        unit="GHz",
        start_frequency=fmin_GHz,
        stop_frequency=fmax_GHz,
        sweep_type="Interpolating",
        num_of_freq_points=101
    )

    hfss.save_project()

    if solve_after:
        hfss.analyze(setup.name)     # resolve Setup1 + sweep
        ensure_amp_phase_vars(hfss, port_names)
        ok = edit_sources_robusto(hfss, port_names)
        if not ok:
            print("[WARN] Edit Sources falhou (chaves). Ajuste manual em Solve > Edit Sources no HFSS.")
        # opcional: criar um gráfico de S(1,1)
        try:
            hfss.post.create_report(expressions=["dB(S(1,1))"], primary_sweep_variable="Freq")
        except Exception:
            pass
        hfss.save_project()

    info = {
        "project_path": project_path,
        "f0_GHz": f0,
        "W_mm": W,
        "L_mm": L,
        "eps_eff": eps_eff,
        "nx": nx,
        "ny": ny,
        "N": n_real,
        "ports": port_names,
        "setup": setup.name
    }
    return hfss, info

# =========================
# GUI
# =========================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HFSS Patch Array – Automação (PyAEDT)")
        self.geometry("720x520")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        # ---- Labels e Entradas ----
        self.grid_columnconfigure(1, weight=1)

        self.lbl_fmin = ctk.CTkLabel(self, text="Frequência mínima (GHz):")
        self.ent_fmin = ctk.CTkEntry(self); self.ent_fmin.insert(0, "2.3")

        self.lbl_fmax = ctk.CTkLabel(self, text="Frequência máxima (GHz):")
        self.ent_fmax = ctk.CTkEntry(self); self.ent_fmax.insert(0, "2.5")

        self.lbl_gain = ctk.CTkLabel(self, text="Ganho alvo do array (dBi):")
        self.ent_gain = ctk.CTkEntry(self); self.ent_gain.insert(0, "12")

        self.lbl_epsr = ctk.CTkLabel(self, text="Permissividade relativa εr:")
        self.ent_epsr = ctk.CTkEntry(self); self.ent_epsr.insert(0, "4.4")

        self.lbl_h = ctk.CTkLabel(self, text="Altura do substrato h (mm):")
        self.ent_h = ctk.CTkEntry(self); self.ent_h.insert(0, "1.57")

        self.var_solve = ctk.BooleanVar(value=True)
        self.chk_solve = ctk.CTkCheckBox(self, text="Rodar simulação após criar", variable=self.var_solve)

        self.btn = ctk.CTkButton(self, text="Criar Array no HFSS", command=self.on_create)

        self.logbox = ctk.CTkTextbox(self, width=680, height=300)

        # ---- Layout ----
        self.lbl_fmin.grid(row=0, column=0, padx=10, pady=(12,4), sticky="e")
        self.ent_fmin.grid(row=0, column=1, padx=10, pady=(12,4), sticky="ew")
        self.lbl_fmax.grid(row=1, column=0, padx=10, pady=4, sticky="e")
        self.ent_fmax.grid(row=1, column=1, padx=10, pady=4, sticky="ew")
        self.lbl_gain.grid(row=2, column=0, padx=10, pady=4, sticky="e")
        self.ent_gain.grid(row=2, column=1, padx=10, pady=4, sticky="ew")
        self.lbl_epsr.grid(row=3, column=0, padx=10, pady=4, sticky="e")
        self.ent_epsr.grid(row=3, column=1, padx=10, pady=4, sticky="ew")
        self.lbl_h.grid(row=4, column=0, padx=10, pady=4, sticky="e")
        self.ent_h.grid(row=4, column=1, padx=10, pady=4, sticky="ew")
        self.chk_solve.grid(row=5, column=1, padx=10, pady=(8,12), sticky="w")
        self.btn.grid(row=6, column=1, padx=10, pady=6, sticky="w")
        self.logbox.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_rowconfigure(7, weight=1)

        self.log("[Dica] Campos com unidades: GHz, dBi, εr, mm.")

    def log(self, msg: str):
        self.logbox.insert("end", msg + "\n")
        self.logbox.see("end")

    def on_create(self):
        try:
            fmin = float(self.ent_fmin.get())
            fmax = float(self.ent_fmax.get())
            if fmax <= fmin:
                self.log("Erro: fmax deve ser maior que fmin.")
                return
            gtar = float(self.ent_gain.get())
            epsr = float(self.ent_epsr.get())
            h = float(self.ent_h.get())

            f0 = 0.5 * (fmin + fmax)
            W, L, eps_eff = hammerstad_patch_dims(f0, epsr, h)
            nx, ny, nreal = estimate_array_layout(gtar)
            self.log(f"[Analítico] f0={f0:.3f} GHz | W≈{W:.2f} mm, L≈{L:.2f} mm, eps_eff≈{eps_eff:.4f}")
            self.log(f"[Array] G_target={gtar:.2f} dBi | N_req≈{nx*ny} | Layout {nx}x{ny} (N_real={nreal})")

            script_dir = os.path.dirname(os.path.abspath(__file__))
            out_dir = os.path.join(script_dir, "examples")
            os.makedirs(out_dir, exist_ok=True)

            hfss, info = create_patch_array_hfss(
                fmin_GHz=fmin, fmax_GHz=fmax,
                g_target_dbi=gtar, eps_r=epsr, h_mm=h,
                solve_after=self.var_solve.get(),
                out_dir=out_dir
            )
            self.log(f"[Projeto] Salvo em: {info['project_path']}")
            self.log(f"[Setup] {info['setup']}")
            self.log(f"[Ports] {', '.join(info['ports'])}")

        except Exception as e:
            self.log("Erro: " + str(e))
            self.log(traceback.format_exc())

# =========================
# Execução
# =========================
if __name__ == "__main__":
    app = App()
    app.mainloop()
