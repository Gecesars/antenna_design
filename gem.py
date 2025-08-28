# -*- coding: utf-8 -*-
import os
import sys
import math
import shutil
import traceback
import customtkinter as ctk
from ansys.aedt.core import Hfss

# ===================== Configurações Gerais =====================
AEDT_VERSION = "2024.2"
UNITS = "mm"
COPPER_T = 0.035
PATCH_GAIN_DBI = 6.5  # Ganho típico de um único patch
Z0_PORT = 50.0
ZPATCH_EDGE = 200.0  # Impedância de borda estimada para o patch
PAD_MIN = 1.0

SETUP_NAME = "MainSetup"
SWEEP_NAME = "FrequencySweep"

# ===================== Utilidades Eletromag =====================
def c_mm_per_GHz() -> float:
    return 299.792458

def hammerstad_patch_dims(f0_GHz, eps_r, h_mm):
    c = c_mm_per_GHz()
    W = c / (2.0 * f0_GHz) * math.sqrt(2.0 / (eps_r + 1.0))
    eps_eff = (eps_r + 1.0) / 2.0 + (eps_r - 1.0) / 2.0 * (1.0 / math.sqrt(1.0 + 12.0 * h_mm / W))
    dL = 0.412 * h_mm * ((eps_eff + 0.3) * (W/h_mm + 0.264)) / ((eps_eff - 0.258) * (W/h_mm + 0.8))
    L_eff = c / (2.0 * f0_GHz * math.sqrt(eps_eff))
    L = L_eff - 2.0 * dL
    return W, L, eps_eff

def _eps_eff_line(eps_r, w_h):
    return (eps_r + 1)/2 + (eps_r - 1)/2 * (1 + 12/w_h) ** -0.5

def z0_from_w_h(eps_r, w_h):
    ee = _eps_eff_line(eps_r, w_h)
    if w_h <= 1:
        return (60 / math.sqrt(ee)) * math.log(8.0 / w_h + 0.25 * w_h)
    else:
        return (120 * math.pi) / (math.sqrt(ee) * (w_h + 1.393 + 0.667 * math.log(w_h + 1.444)))

def w_for_z0(eps_r, h_mm, z0_ohm):
    lo, hi, tolerance, max_iter = 0.05, 20.0, 0.001, 50
    for _ in range(max_iter):
        mid = (lo + hi) / 2.0; z0_calc = z0_from_w_h(eps_r, mid)
        if abs(z0_calc - z0_ohm) < tolerance: break
        if z0_calc > z0_ohm: lo = mid
        else: hi = mid
    return mid * h_mm

def quarter_wave_len_mm(f0_GHz, eps_r, w_mm, h_mm):
    w_h = max(w_mm / h_mm, 1e-6); ee = _eps_eff_line(eps_r, w_h)
    lambda_g = c_mm_per_GHz() / (f0_GHz * math.sqrt(ee))
    return lambda_g / 4.0

# ===================== Arquivos / Projeto ======================
def clean_previous(project_path: str):
    # (Função de limpeza mantida como estava)
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

# ===================== Construção do Modelo =====================
def create_param(hfss: Hfss, name: str, expr: str):
    hfss[name] = expr

def assign_lumped_port_with_integration_line(hfss: Hfss, sheet_name: str, port_name: str, start_coord: list, end_coord: list):
    """Cria uma porta lumped em um sheet, definindo explicitamente a linha de integração."""
    try:
        # A referência ao GND é implícita quando a linha de integração é fornecida
        # e o sheet toca fisicamente o plano de terra.
        port = hfss.lumped_port(
            assignment=sheet_name,
            impedance=Z0_PORT,
            renormalize=True,
            name=port_name,
            integration_line_start=start_coord,
            integration_line_stop=end_coord
        )
        return port
    except Exception as e:
        print(f"Erro ao criar porta {port_name}: {str(e)}")
        raise

def build_array_project(
    fmin_GHz: float, fmax_GHz: float, g_target_dbi: float,
    eps_r: float, h_mm: float, out_dir: str, solve_after: bool = True,
) -> tuple[Hfss, dict]:
    os.makedirs(out_dir, exist_ok=True)
    project_path = os.path.join(out_dir, "patch_array_hfss_modal.aedt")
    clean_previous(project_path)

    f0 = 0.5 * (fmin_GHz + fmax_GHz)
    fstart = fmin_GHz; fstop = fmax_GHz

    Wp, Lp, eps_eff_patch = hammerstad_patch_dims(f0, eps_r, h_mm)

    # Lógica para determinar o número de elementos do arranjo
    gain_factor_needed = 10 ** ((g_target_dbi - PATCH_GAIN_DBI) / 10.0)
    n_elements = max(1, math.ceil(gain_factor_needed))
    nx = int(round(math.sqrt(n_elements)))
    ny = int(math.ceil(n_elements / nx))
    N = nx * ny

    Zt = math.sqrt(Z0_PORT * ZPATCH_EDGE); Wt = w_for_z0(eps_r, h_mm, Zt)
    Lq = quarter_wave_len_mm(f0, eps_r, Wt, h_mm); Pad = max(PAD_MIN, 0.6 * Wt)
    
    lam0 = c_mm_per_GHz() / f0; pitch = 0.75 * lam0  # Espaçamento típico entre 0.5 e 0.9 lambda
    sx = (nx - 1) * pitch + Wp + 20; sy = (ny - 1) * pitch + Lp + 2 * (Lq + Pad) + 20

    with Hfss(
        projectname=project_path, designname="PatchArray_HFSS", solution_type="Modal",
        specified_version=AEDT_VERSION, non_graphical=False, new_desktop_session=True) as hfss:
        
        hfss.modeler.model_units = UNITS
        z_top = f"{h_mm+COPPER_T:.6f}mm"

        gnd = hfss.modeler.create_box(origin=[f"{-sx/2}", f"{-sy/2}", "0"], sizes=[f"{sx}", f"{sy}", f"{COPPER_T}"], name="GND", matname="pec")
        sub = hfss.modeler.create_box(origin=[f"{-sx/2}", f"{-sy/2}", f"{COPPER_T}"], sizes=[f"{sx}", f"{sy}", f"{h_mm}"], name="SUB", matname="FR4_epoxy")
        
        all_copper_parts = []
        port_names: List[str] = []
        x0 = -(nx - 1) * pitch / 2.0
        y0 = -(ny - 1) * pitch / 2.0

        for ix in range(nx):
            for iy in range(ny):
                el_name = f"_{ix+1}_{iy+1}"
                cx = x0 + ix * pitch
                cy = y0 + iy * pitch

                patch = hfss.modeler.create_box(origin=[f"{cx-Wp/2}", f"{cy-Lp/2}", z_top], sizes=[f"{Wp}", f"{Lp}", f"{COPPER_T}"], name=f"Patch{el_name}", matname="pec")
                
                patch_y_min = cy - Lp/2
                line = hfss.modeler.create_box(origin=[f"{cx-Wt/2}", f"{patch_y_min-Lq}", z_top], sizes=[f"{Wt}", f"{Lq}", f"{COPPER_T}"], name=f"Tpline{el_name}", matname="pec")
                
                pad_y_start = patch_y_min - Lq
                pad = hfss.modeler.create_box(origin=[f"{cx-Wt/2}", f"{pad_y_start-Pad}", z_top], sizes=[f"{Wt}", f"{Pad}", f"{COPPER_T}"], name=f"Feedpad{el_name}", matname="pec")
                
                all_copper_parts.extend([patch, line, pad])
                
                # --- Criação da Porta Lumped com Linha de Integração ---
                sheet_name = f"PortSheet{el_name}"
                port_y_pos = pad_y_start - Pad
                
                port_sheet = hfss.modeler.create_rectangle(
                    origin=[f"{cx-Wt/2}", f"{port_y_pos}", "0"],
                    sizes=[f"{Wt}", f"{h_mm + COPPER_T}"],
                    orientation="XZ",
                    name=sheet_name
                )

                port_name = f"P{len(port_names)+1}"
                # Coordenadas da linha de integração: do GND ao centro do feed
                start_coord = [f"{cx}mm", f"{port_y_pos}mm", "0mm"]
                end_coord = [f"{cx}mm", f"{port_y_pos}mm", f"{h_mm+COPPER_T}mm"]
                
                assign_lumped_port_with_integration_line(hfss, sheet_name, port_name, start_coord, end_coord)
                port_names.append(port_name)

        hfss.modeler.unite(all_copper_parts, keep_originals=False)
        hfss.assign_perfecte_to_sheets("GND")
        
        hfss.create_open_region(frequency=f"{f0}GHz")

        setup = hfss.create_setup(setupname=SETUP_NAME)
        setup.props["Frequency"] = f"{f0}GHz"; setup.props["MaximumPasses"] = 6
        setup.props["MaximumDeltaS"] = 0.02; setup.update()
        
        setup.create_frequency_sweep(
            sweepname=SWEEP_NAME, unit="GHz", freqstart=fstart, freqstop=fstop,
            sweep_type="Interpolating", num_of_freq_points=101
        )
        hfss.save_project()
        
        if solve_after:
            hfss.analyze_setup(SETUP_NAME)
            hfss.post.create_report([f"dB(S({p},{p}))" for p in port_names])
            hfss.post.create_far_fields_report(expressions="GainTotal", plot_type="3D Polar Plot")
            hfss.save_project()

        info = { "project_path": project_path, "f0_GHz": f0, "Wp_mm": Wp, "Lp_mm": Lp, "eps_eff_patch": eps_eff_patch,
                 "Zt_ohm": Zt, "Wt_mm": Wt, "Lq_mm": Lq, "Pad_mm": Pad, "nx": nx, "ny": ny, "N": N, "ports": port_names,
                 "setup": SETUP_NAME, "sweep": SWEEP_NAME, "fstart_GHz": fstart, "fstop_GHz": fstop }
        return hfss, info

# ===================== GUI e Main =====================
# (A classe App da GUI e o bloco __main__ permanecem os mesmos, omitidos para brevidade)
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HFSS Patch Array Designer"); self.geometry("820x600")
        ctk.set_appearance_mode("dark"); ctk.set_default_color_theme("dark-blue")
        self.hfss_ref: Hfss | None = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._mk_row(0, "Frequência mínima (GHz):", "2.3")
        self._mk_row(1, "Frequência máxima (GHz):", "2.5")
        self._mk_row(2, "Ganho alvo do array (dBi):", "8")
        self._mk_row(3, "εr (FR4≈4.4):", "4.4")
        self._mk_row(4, "Altura do substrato h (mm):", "1.57")
        self.chk_run = ctk.CTkCheckBox(self, text="Rodar simulação após criar"); self.chk_run.grid(row=5, column=1, padx=10, pady=(6, 6), sticky="w"); self.chk_run.select()
        self.btn = ctk.CTkButton(self, text="Criar e Simular Array no HFSS", command=self.on_create); self.btn.grid(row=6, column=1, padx=10, pady=(0, 8), sticky="w")
        self.txt = ctk.CTkTextbox(self, width=780, height=360); self.txt.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_columnconfigure(1, weight=1); self.grid_rowconfigure(7, weight=1)

    def _mk_row(self, r, label, default):
        ctk.CTkLabel(self, text=label).grid(row=r, column=0, padx=10, pady=6, sticky="e")
        e = ctk.CTkEntry(self); e.insert(0, default); e.grid(row=r, column=1, padx=10, pady=6, sticky="ew")
        setattr(self, f"e{r}", e)

    def _log(self, s: str):
        self.txt.insert("end", s + "\n"); self.txt.see("end")

    def on_create(self):
        try:
            fmin = float(self.e0.get()); fmax = float(self.e1.get())
            if fmax <= fmin: raise ValueError("fmax deve ser maior que fmin.")
            gtar = float(self.e2.get()); epsr = float(self.e3.get()); h = float(self.e4.get())
            run_after = self.chk_run.get()
            f0 = 0.5*(fmin+fmax); Wp, Lp, ee = hammerstad_patch_dims(f0, epsr, h)
            self._log(f"[Analítico] f0={f0:.3f} GHz | W≈{Wp:.2f} mm, L≈{Lp:.2f} mm, εeff≈{ee:.4f}")
            out_dir = os.path.dirname(os.path.abspath(__file__))
            hfss, info = build_array_project(
                fmin_GHz=fmin, fmax_GHz=fmax, g_target_dbi=gtar,
                eps_r=epsr, h_mm=h, out_dir=out_dir, solve_after=run_after
            )
            self.hfss_ref = hfss
            self._log(f"[Projeto] {info['project_path']}"); self._log(f"[Linha λ/4] Zt≈{info['Zt_ohm']:.1f} Ω | Wt≈{info['Wt_mm']:.2f} mm | Lq≈{info['Lq_mm']:.2f} mm | Pad≈{info['Pad_mm']:.2f} mm")
            self._log(f"[Sweep] {info['setup']} : {info['sweep']}  ({info['fstart_GHz']:.3f}–{info['fstop_GHz']:.3f} GHz)"); self._log(f"[Ports] {', '.join(info['ports'])}")
            self._log(f"[Array] {info['nx']}×{info['ny']} = {info['N']} elementos")
            if run_after: self._log("[Simulação] Concluída. Verifique os resultados no HFSS e os arquivos exportados.")
            else: self._log("[Simulação] Modelo criado. A análise não foi executada.")
        except Exception as e:
            self._log("ERRO: " + str(e)); self._log(traceback.format_exc())

    def on_close(self):
        try:
            if self.hfss_ref is not None:
                self.hfss_ref.release_desktop(close_projects=False, close_desktop=True)
                self._log("[AEDT] Instância liberada.")
        except Exception as e:
            self._log(f"[AEDT] Falha ao liberar: {e}")
        finally:
            self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()