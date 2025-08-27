# -*- coding: utf-8 -*-
"""
Patch Array com λ/4 (PyAEDT 0.18.1) + CTk
- Porta LUMPED correta (sheet XZ) no fim da linha 50 Ω: largura=W50, altura=h+tCu, tocando GND (z=0) e condutor (z=h).
- Setup "OpSetup" + sweep "OpFast" (Interpolating, 101 pts, banda expandida em 30% no total) e execução explícita.
- UI em CustomTkinter. AEDT é liberado ao fechar a janela.

Requisitos:
  pip install ansys-aedt-core==0.18.1 customtkinter
"""

from __future__ import annotations
import os, sys, math, traceback
from dataclasses import dataclass
from typing import List, Tuple

# ---- constantes físicas (scikit-rf opcional) ----
try:
    from skrf.constants import c as C0
except Exception:
    C0 = 299_792_458.0

# ---- PyAEDT ----
from ansys.aedt.core import Hfss

# ---- GUI ----
import customtkinter as ctk


# ===================== utilidades analíticas =====================

def mm(x: float) -> str:
    return f"{x:.6f}mm"

def patch_dimensions(f0_ghz: float, eps_r: float, h_mm: float) -> Tuple[float, float, float]:
    f0 = f0_ghz * 1e9
    h = h_mm / 1000.0
    W = (C0 / (2.0 * f0)) * math.sqrt(2.0 / (eps_r + 1.0))
    eps_eff = 0.5*(eps_r+1) + 0.5*(eps_r-1)*(1+12*h/W)**-0.5
    dL_over_h = 0.412*((eps_eff+0.3)*(W/h+0.264))/((eps_eff-0.258)*(W/h+0.8))
    L_eff = C0/(2.0*f0*math.sqrt(eps_eff))
    L = L_eff - 2.0*dL_over_h*h
    return W*1000.0, L*1000.0, eps_eff

def rin_edge_patch(f0_ghz: float, W_mm: float, h_mm: float) -> float:
    f0 = f0_ghz*1e9
    lam0 = C0/f0
    W = W_mm/1000.0
    h = h_mm/1000.0
    k0h = 2*math.pi*h/lam0
    G = math.pi*W/(120.0*math.pi*lam0)*(1.0 - (k0h**2)/24.0)
    return 1.0/(2.0*G + 1e-18)

def microstrip_w_for_z0(z0: float, eps_r: float, h_mm: float) -> float:
    """Hammerstad-Jensen aproximado."""
    h = h_mm/1000.0
    A = z0/60.0*math.sqrt((eps_r+1.0)/2.0) + ((eps_r-1.0)/(eps_r+1.0))*(0.23+0.11/eps_r)
    B = (377*math.pi)/(2.0*z0*math.sqrt(eps_r))
    if A <= 1:
        W_h = 8*math.exp(A)/(math.exp(2*A)-2)
    else:
        W_h = (2/math.pi)*(B - 1 - math.log(2*B-1) + ((eps_r-1)/(2*eps_r))*(math.log(B-1)+0.39-0.61/eps_r))
    return max(0.05, W_h*h*1000.0)

def microstrip_eps_eff(w_mm: float, eps_r: float, h_mm: float) -> float:
    w_h = (w_mm/1000.0)/(h_mm/1000.0)
    if w_h <= 1:
        return (eps_r+1)/2 + (eps_r-1)/2*(1/math.sqrt(1+12/w_h) + 0.04*(1-w_h)**2)
    return (eps_r+1)/2 + (eps_r-1)/2*(1/math.sqrt(1+12/w_h))

def guided_quarter_lambda(f0_ghz: float, eps_eff_line: float) -> float:
    f0 = f0_ghz*1e9
    return (C0/(f0*math.sqrt(eps_eff_line)))*1000.0/4.0

def gain_to_n_elements(g_target_dbi: float, g_elem_dbi: float = 6.0) -> int:
    return max(1, int(math.ceil(10**((g_target_dbi - g_elem_dbi)/10.0))))

def square_grid(n: int) -> Tuple[int,int]:
    m = int(math.floor(math.sqrt(n)))
    k = int(math.ceil(n/m))
    return m, k


# ===================== HFSS helpers =====================

@dataclass
class PatchParams:
    Wp: float
    Lp: float
    Wt: float
    Lq: float
    W50: float
    eps_eff: float

def set_vars(hfss: Hfss, h_mm: float, tcu_mm: float, p: PatchParams):
    hfss["h"]   = mm(h_mm)
    hfss["tCu"] = mm(tcu_mm)
    hfss["Wp"]  = mm(p.Wp)
    hfss["Lp"]  = mm(p.Lp)
    hfss["Wt"]  = mm(p.Wt)
    hfss["Lq"]  = mm(p.Lq)
    hfss["W50"] = mm(p.W50)

def add_ground_and_substrate(hfss: Hfss, size_x_mm: float, size_y_mm: float, h_mm: float):
    hfss.modeler.create_box(
        origin=[mm(-size_x_mm/2), mm(-size_y_mm/2), "0mm"],
        sizes=[mm(size_x_mm), mm(size_y_mm), mm(h_mm)],
        name="Substrate", material="FR4_epoxy"
    )
    gnd = hfss.modeler.create_rectangle(
        origin=[mm(-size_x_mm/2), mm(-size_y_mm/2), "0mm"],
        sizes=[mm(size_x_mm), mm(size_y_mm)], orientation="XY", name="GND"
    )
    hfss.assign_perfecte_to_sheets(["GND"])
    return gnd

def add_patch_with_qw(hfss: Hfss, cx_mm: float, cy_mm: float, p: PatchParams, idx: int,
                      h_mm: float, tcu_mm: float, z0: float = 50.0) -> str:
    z_top = "h"  # topo do substrato (condutor em sheet)
    # Patch
    patch = hfss.modeler.create_rectangle(
        origin=[f"{mm(cx_mm)}-Wp/2", f"{mm(cy_mm)}-Lp/2", z_top],
        sizes=["Wp", "Lp"], orientation="XY", name=f"Patch_{idx}"
    )
    # λ/4
    tline = hfss.modeler.create_rectangle(
        origin=[f"{mm(cx_mm)}-Wt/2", f"{mm(cy_mm)}+Lp/2", z_top],
        sizes=["Wt", "Lq"], orientation="XY", name=f"TLineQW_{idx}"
    )
    # 50 Ω
    Lfeed = max(5.0, 0.15*p.Lq)
    feed = hfss.modeler.create_rectangle(
        origin=[f"{mm(cx_mm)}-W50/2", f"{mm(cy_mm)}+Lp/2+Lq", z_top],
        sizes=["W50", mm(Lfeed)], orientation="XY", name=f"Feed50_{idx}"
    )

    hfss.modeler.unite([patch, tline, feed])
    hfss.assign_perfecte_to_sheets([patch.name])  # condutor superior

    # Porta lumped: sheet XZ que CONECTA condutor (z=h) ao GND (z=0)
    y_port = cy_mm + (p.Lp/2.0) + p.Lq + Lfeed
    port_sheet = hfss.modeler.create_rectangle(
        origin=[f"{mm(cx_mm)}-W50/2", mm(y_port), "0mm"],
        sizes=["W50", "h+tCu"], orientation="XZ", name=f"PortSheet_{idx}"
    )

    pname = f"P{idx}"
    # 1) Tentar lumped (correto para microstrip)
    try:
        hfss.lumped_port(assignment=port_sheet, reference=None, impedance=z0,
                         renormalize=True, name=pname)
    except Exception:
        # 2) Fallback: wave port usando a mesma folha e linha de integração vertical
        int_line = [[mm(cx_mm), mm(y_port), "0mm"], [mm(cx_mm), mm(y_port), "h+tCu"]]
        hfss.wave_port(assignment=port_sheet, name=pname, integration_line=int_line)
    return pname


def build_array_project(project_path: str,
                        f_low_ghz: float, f_high_ghz: float, f0_ghz: float,
                        g_target_dbi: float, eps_r: float, h_mm: float, tcu_mm: float,
                        run_after: bool = True, version: str = "2024.2"):
    # ---- dimensões e casamento ----
    Wp, Lp, eps_eff_patch = patch_dimensions(f0_ghz, eps_r, h_mm)
    Rin_edge = rin_edge_patch(f0_ghz, Wp, h_mm)
    Zt = math.sqrt(50.0*Rin_edge)
    Wt = microstrip_w_for_z0(Zt, eps_r, h_mm)
    W50 = microstrip_w_for_z0(50.0, eps_r, h_mm)
    Lq = guided_quarter_lambda(f0_ghz, microstrip_eps_eff(Wt, eps_r, h_mm))
    p = PatchParams(Wp, Lp, Wt, Lq, W50, eps_eff_patch)

    # ---- arranjo ----
    lam0_mm = (C0/(f0_ghz*1e9))*1000.0
    pitch = 0.6*lam0_mm
    Ny, Nx = square_grid(gain_to_n_elements(g_target_dbi, 6.0))
    n_real = Nx*Ny

    margin = 0.5*lam0_mm
    size_x = Nx*pitch + 2*margin
    size_y = Ny*pitch + 2*margin

    # ---- HFSS ----
    hfss = Hfss(project=project_path, design="PatchArray_HFSS",
                solution_type="Modal", new_desktop=True,
                non_graphical=False, version=version)
    hfss.modeler.model_units = "mm"
    set_vars(hfss, h_mm, tcu_mm, p)
    add_ground_and_substrate(hfss, size_x, size_y, h_mm)

    # elementos
    ports: List[str] = []
    x0 = -(Nx-1)*pitch/2.0
    y0 = -(Ny-1)*pitch/2.0
    idx = 1
    for iy in range(Ny):
        for ix in range(Nx):
            cx = x0 + ix*pitch
            cy = y0 + iy*pitch
            ports.append(add_patch_with_qw(hfss, cx, cy, p, idx, h_mm, tcu_mm, 50.0))
            idx += 1

    # região aberta
    hfss.create_open_region(frequency=f"{f0_ghz}GHz")

    # ---- setup explícito + sweep FAST 101 ----
    span = f_high_ghz - f_low_ghz
    f1 = max(0.01, f_low_ghz - 0.15*span)
    f2 = f_high_ghz + 0.15*span

    if "OpSetup" in hfss.setups:
        hfss.delete_setup("OpSetup")
    setup = hfss.create_setup(name="OpSetup", Frequency=f"{f0_ghz}GHz")

    if setup.sweeps and "OpFast" in setup.sweeps:
        setup.delete_sweep("OpFast")
    setup.create_frequency_sweep(name="OpFast", unit="GHz",
                                 start_frequency=f1, stop_frequency=f2,
                                 sweep_type="Interpolating", num_of_freq_points=101)

    hfss.active_setup = "OpSetup"
    hfss.save_project()

    # ---- execução + fontes pós-processamento (opcional) ----
    if run_after:
        # garantir bloqueio até terminar
        try:
            hfss.analyze_setup("OpSetup")
        except Exception:
            try:
                hfss.analyze(setup="OpSetup")
            except Exception:
                hfss.analyze()

        # aplicar excitações de pós-processamento somente se a solução existir
        try:
            assign = {}
            pwr_each = 1.0/max(1, len(ports))
            for p_name in ports:
                hfss[f"Amp_{p_name}"] = f"{pwr_each:.6f}W"
                hfss[f"Phi_{p_name}"] = "0deg"
                assign[p_name] = {"magnitude": f"Amp_{p_name}",
                                  "phase": f"Phi_{p_name}",
                                  "source_type": "Power"}
            try:
                hfss.edit_sources(assign)
            except Exception:
                # alguns builds exigem ':1'
                assign_modes = {f"{k}:1": v for k, v in assign.items()}
                hfss.edit_sources(assign_modes)
        except Exception:
            # Se não der, seguimos sem editar fontes; S-params continuam válidos.
            pass

        # relatório simples
        try:
            hfss.post.create_report(
                expressions=["dB(S(1,1))"],
                setup_sweep_name="OpSetup : OpFast",
                primary_sweep_variable="Freq",
            )
        except Exception:
            pass

    info = {
        "W_mm": Wp, "L_mm": Lp, "eps_eff": eps_eff_patch,
        "Rin_edge_ohm": Rin_edge, "Zt_ohm": Zt,
        "Wt_mm": Wt, "Lq_mm": Lq, "W50_mm": W50,
        "Nx": Nx, "Ny": Ny, "N_real": n_real,
        "ports": ports, "setup": "OpSetup", "sweep": "OpFast",
        "f1_ghz": f1, "f2_ghz": f2
    }
    return hfss, info


# ===================== GUI (CTk) =====================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Patch Array λ/4 – HFSS (PyAEDT)")
        self.geometry("760x560")
        ctk.set_appearance_mode("dark"); ctk.set_default_color_theme("dark-blue")
        self.hfss: Hfss|None = None

        r = 0
        ctk.CTkLabel(self, text="[Dica] Unidades: GHz, dBi, εr, mm.").grid(row=r, column=0, columnspan=4, sticky="w", padx=10, pady=(10,6)); r+=1

        ctk.CTkLabel(self, text="F_low (GHz)").grid(row=r, column=0, sticky="e")
        self.e_flow = ctk.CTkEntry(self); self.e_flow.insert(0,"2.30"); self.e_flow.grid(row=r, column=1, sticky="w", padx=6, pady=6)
        ctk.CTkLabel(self, text="F_high (GHz)").grid(row=r, column=2, sticky="e")
        self.e_fhigh = ctk.CTkEntry(self); self.e_fhigh.insert(0,"2.50"); self.e_fhigh.grid(row=r, column=3, sticky="w", padx=6, pady=6); r+=1

        ctk.CTkLabel(self, text="f0 (GHz)").grid(row=r, column=0, sticky="e")
        self.e_f0 = ctk.CTkEntry(self); self.e_f0.insert(0,"2.40"); self.e_f0.grid(row=r, column=1, sticky="w", padx=6, pady=6)
        ctk.CTkLabel(self, text="Ganho alvo (dBi)").grid(row=r, column=2, sticky="e")
        self.e_gain = ctk.CTkEntry(self); self.e_gain.insert(0,"12.0"); self.e_gain.grid(row=r, column=3, sticky="w", padx=6, pady=6); r+=1

        ctk.CTkLabel(self, text="εr").grid(row=r, column=0, sticky="e")
        self.e_eps = ctk.CTkEntry(self); self.e_eps.insert(0,"4.4"); self.e_eps.grid(row=r, column=1, sticky="w", padx=6, pady=6)
        ctk.CTkLabel(self, text="h (mm)").grid(row=r, column=2, sticky="e")
        self.e_h = ctk.CTkEntry(self); self.e_h.insert(0,"1.57"); self.e_h.grid(row=r, column=3, sticky="w", padx=6, pady=6); r+=1

        ctk.CTkLabel(self, text="tCu (mm)").grid(row=r, column=0, sticky="e")
        self.e_tcu = ctk.CTkEntry(self); self.e_tcu.insert(0,"0.035"); self.e_tcu.grid(row=r, column=1, sticky="w", padx=6, pady=6); r+=1

        self.var_run = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(self, text="Rodar simulação após criar", variable=self.var_run).grid(row=r, column=0, columnspan=2, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkButton(self, text="Criar Array no HFSS", command=self.on_create).grid(row=r, column=0, columnspan=2, sticky="w", padx=10, pady=10); r+=1

        self.txt = ctk.CTkTextbox(self, height=260)
        self.txt.grid(row=r, column=0, columnspan=4, sticky="nsew", padx=10, pady=10)
        self.grid_rowconfigure(r, weight=1)
        for col in range(4):
            self.grid_columnconfigure(col, weight=1)

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def log(self, s:str):
        self.txt.insert("end", s+"\n"); self.txt.see("end"); print(s, flush=True)

    def on_create(self):
        self.txt.delete("1.0","end")
        try:
            f_low  = float(self.e_flow.get().replace(',','.'))
            f_high = float(self.e_fhigh.get().replace(',','.'))
            f0     = float(self.e_f0.get().replace(',','.'))
            g      = float(self.e_gain.get().replace(',','.'))
            eps    = float(self.e_eps.get().replace(',','.'))
            h      = float(self.e_h.get().replace(',','.'))
            tcu    = float(self.e_tcu.get().replace(',','.'))
        except Exception:
            self.log("Parâmetros inválidos."); return

        self.log(f"[Analítico] f0={f0:.3f} GHz")
        Wp, Lp, eps_eff = patch_dimensions(f0, eps, h)
        self.log(f"[Patch] W≈{Wp:.2f} mm, L≈{Lp:.2f} mm, εeff≈{eps_eff:.4f}")

        try:
            base = os.path.dirname(__file__)
        except NameError:
            base = os.getcwd()
        proj = os.path.join(base, "PatchArray_HFSS.aedt")

        try:
            self.hfss, info = build_array_project(
                project_path=proj,
                f_low_ghz=f_low, f_high_ghz=f_high, f0_ghz=f0,
                g_target_dbi=g, eps_r=eps, h_mm=h, tcu_mm=tcu,
                run_after=self.var_run.get(), version="2024.2"
            )
        except Exception:
            self.log("Erro ao criar o projeto:")
            self.log(traceback.format_exc()); return

        self.log(f"[Array] alvo={g:.2f} dBi | Nx={info['Nx']} Ny={info['Ny']} (N={info['N_real']})")
        self.log(f"[λ/4] Wt≈{info['Wt_mm']:.2f} mm | Lq≈{info['Lq_mm']:.2f} mm | W50≈{info['W50_mm']:.2f} mm")
        self.log(f"[Setup] Executado: {info['setup']} : {info['sweep']}  ({info['f1_ghz']:.3f}–{info['f2_ghz']:.3f} GHz, FAST 101 pts)")
        self.log(f"[Ports] {', '.join(info['ports'])}")
        self.log(f"[Projeto] {proj}")

    def on_close(self):
        try:
            if self.hfss is not None:
                self.hfss.save_project()
                self.hfss.release_desktop()
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
