# -*- coding: utf-8 -*-
import os
import sys
import math
import shutil
import traceback
import customtkinter as ctk
from ansys.aedt.core import Hfss
from ansys.aedt.core.modeler.advanced_cad.stackup_3d import Stackup3D

# ===================== Configurações Gerais =====================
AEDT_VERSION = "2024.2"
UNITS = "mm"
COPPER_T = 0.035
PATCH_GAIN_DBI = 6.5  # Ganho típico de um único patch para cálculo do array

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

# ===================== Arquivos / Projeto ======================
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

# ===================== Construção do Modelo =====================
def build_array_project(
    fmin_GHz: float, fmax_GHz: float, g_target_dbi: float,
    eps_r: float, h_mm: float, spacing_factor: float, feed_offset_mm: float,
    out_dir: str, solve_after: bool = True,
) -> tuple[Hfss, dict]:
    os.makedirs(out_dir, exist_ok=True)
    project_path = os.path.join(out_dir, "patch_array_stackup.aedt")
    clean_previous(project_path)

    f0 = 0.5 * (fmin_GHz + fmax_GHz)
    fstart = fmin_GHz; fstop = fmax_GHz

    Wp, Lp, eps_eff_patch = hammerstad_patch_dims(f0, eps_r, h_mm)

    gain_factor_needed = 10 ** ((g_target_dbi - PATCH_GAIN_DBI) / 10.0)
    n_elements = max(1, math.ceil(gain_factor_needed))
    nx = int(round(math.sqrt(n_elements)))
    ny = int(math.ceil(n_elements / nx))
    N = nx * ny

    lam0 = c_mm_per_GHz() / f0
    pitch = spacing_factor * lam0
    
    total_width = (nx - 1) * pitch + Wp + pitch
    total_length = (ny - 1) * pitch + Lp + pitch

    with Hfss(
        projectname=project_path, designname="PatchArray_HFSS_Stackup", solution_type="Terminal",
        specified_version=AEDT_VERSION, non_graphical=False, new_desktop_session=True) as hfss:
        
        hfss.modeler.model_units = UNITS
        
        # --- Criação da Estrutura em Camadas (Stackup) ---
        stackup = Stackup3D(hfss)
        
        ground = stackup.add_ground_layer("Ground", material="pec", thickness=COPPER_T)
        stackup.add_dielectric_layer("Substrate", thickness=h_mm, material="FR4_epoxy")
        signal = stackup.add_signal_layer("Signal", material="pec", thickness=COPPER_T)

        stackup.dielectric_x_position = f"{-total_width/2}mm"
        stackup.dielectric_y_position = f"{-total_length/2}mm"
        stackup.dielectric_x_size = f"{total_width}mm"
        stackup.dielectric_y_size = f"{total_length}mm"

        # --- Criação do Arranjo de Patches e Portas (Loop) ---
        start_x = -(nx - 1) * pitch / 2.0
        start_y = -(ny - 1) * pitch / 2.0
        
        port_names = []
        for ix in range(nx):
            for iy in range(ny):
                cx = start_x + ix * pitch
                cy = start_y + iy * pitch
                patch_name = f"Patch_{ix+1}_{iy+1}"
                
                patch = signal.add_patch(
                    patch_width=Wp, patch_length=Lp,
                    patch_name=patch_name, center_position=[cx, cy]
                )
                
                # Cria a alimentação coaxial (probe feed) para este patch
                # Esta função cria o pino, o furo no terra e a Wave Port automaticamente
                patch.create_probe_port(ground, x_offset=feed_offset_mm)
                port_names.append(patch_name)

        # --- Contornos, Setup e Análise ---
        region = hfss.modeler.create_region(pad_percent=300)
        hfss.assign_radiation_boundary_to_objects(region)

        setup = hfss.create_setup(setupname="MainSetup")
        setup.props["Frequency"] = f"{f0}GHz"
        setup.props["MaximumPasses"] = 10
        setup.props["MaximumDeltaS"] = 0.02
        setup.update()

        setup.create_frequency_sweep(
            sweepname="FrequencySweep", unit="GHz", freqstart=fstart, freqstop=fstop,
            sweep_type="Interpolating",
        )
        hfss.save_project()
        
        if solve_after:
            hfss.analyze_setup("MainSetup")
            s_params = [f"S({p},{p})" for p in port_names]
            hfss.post.create_report(expressions=[f"db({s})" for s in s_params])
            hfss.post.create_far_fields_report(expressions="GainTotal", plot_type="3D Polar Plot")
            hfss.save_project()

        info = { "project_path": project_path, "f0_GHz": f0, "Wp_mm": Wp, "Lp_mm": Lp, 
                 "eps_eff_patch": eps_eff_patch, "nx": nx, "ny": ny, "N": N, "ports": port_names,
                 "setup": "MainSetup", "sweep": "FrequencySweep", "fstart_GHz": fstart, "fstop_GHz": fstop }
        return hfss, info

# ===================== GUI e Main =====================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HFSS Patch Array Designer (Stackup3D)"); self.geometry("820x600")
        ctk.set_appearance_mode("dark"); ctk.set_default_color_theme("dark-blue")
        self.hfss_ref: Hfss | None = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self._mk_row(0, "Frequência Mínima (GHz):", "2.3")
        self._mk_row(1, "Frequência Máxima (GHz):", "2.5")
        self._mk_row(2, "Ganho Alvo do Array (dBi):", "10")
        self._mk_row(3, "εr do Substrato (FR4≈4.4):", "4.4")
        self._mk_row(4, "Altura do Substrato h (mm):", "1.57")
        self._mk_row(5, "Espaçamento (em λ₀, ex: 0.75):", "0.75")
        self._mk_row(6, "Deslocamento da Alimentação (mm):", "4.0")

        self.chk_run = ctk.CTkCheckBox(self, text="Rodar simulação após criar"); self.chk_run.grid(row=7, column=1, padx=10, pady=(10, 6), sticky="w"); self.chk_run.select()
        self.btn = ctk.CTkButton(self, text="Criar e Simular Array no HFSS", command=self.on_create); self.btn.grid(row=8, column=1, padx=10, pady=(0, 8), sticky="w")
        self.txt = ctk.CTkTextbox(self, width=780, height=300); self.txt.grid(row=9, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_columnconfigure(1, weight=1); self.grid_rowconfigure(9, weight=1)

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
            spacing = float(self.e5.get()); feed_offset = float(self.e6.get())
            run_after = self.chk_run.get()
            
            f0 = 0.5*(fmin+fmax); Wp, Lp, ee = hammerstad_patch_dims(f0, epsr, h)
            self._log(f"[Analítico] f0={f0:.3f} GHz | W≈{Wp:.2f} mm, L≈{Lp:.2f} mm, εeff≈{ee:.4f}")
            out_dir = os.path.dirname(os.path.abspath(__file__))
            
            self.btn.configure(state="disabled", text="Processando no HFSS..."); self.update_idletasks()
            
            hfss, info = build_array_project(
                fmin_GHz=fmin, fmax_GHz=fmax, g_target_dbi=gtar,
                eps_r=epsr, h_mm=h, spacing_factor=spacing, feed_offset_mm=feed_offset,
                out_dir=out_dir, solve_after=run_after
            )
            self.hfss_ref = hfss
            
            self._log(f"[Projeto] {info['project_path']}"); 
            self._log(f"[Sweep] {info['setup']} : {info['sweep']}  ({info['fstart_GHz']:.3f}–{info['fstop_GHz']:.3f} GHz)")
            self._log(f"[Ports] {', '.join(info['ports'])}"); self._log(f"[Array] {info['nx']}×{info['ny']} = {info['N']} elementos")
            if run_after: self._log("[Simulação] Concluída. Verifique os resultados no HFSS.")
            else: self._log("[Simulação] Modelo criado. A análise não foi executada.")
        except Exception as e:
            self._log("ERRO: " + str(e)); self._log(traceback.format_exc())
        finally:
            self.btn.configure(state="normal", text="Criar e Simular Array no HFSS")

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