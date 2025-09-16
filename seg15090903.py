# -*- coding: utf-8 -*-
import os
import tempfile
from datetime import datetime
import math
import json
import traceback
import queue
import threading
from typing import Tuple, List, Optional

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk

from ansys.aedt.core import Desktop, Hfss

# ---------- Aparência ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ModernPatchAntennaDesigner:
    def __init__(self):
        # AEDT
        self.desktop: Optional[Desktop] = None
        self.hfss: Optional[Hfss] = None
        self.temp_folder = None
        self.project_path = ""
        self.project_title = "patch_array"
        self.design_name = "patch_array"

        # GUI/log
        self.log_queue = queue.Queue()
        self.is_simulation_running = False
        self.save_project = False
        self.stop_simulation = False
        self.simulation_data = None

        # Estados
        self.optimized = False
        self._shape_id = 0  # contador para nomes

        # ---------- Parâmetros do usuário ----------
        self.params = {
            "frequency": 10.0,             # GHz
            "gain": 12.0,                  # dBi
            "sweep_start": 8.0,            # GHz
            "sweep_stop": 12.0,            # GHz
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,

            "spacing_type": "lambda/2",
            "substrate_material": "Duroid (tm)",
            "substrate_thickness": 0.5,    # mm
            "metal_thickness": 0.035,      # mm
            "er": 2.2,
            "tan_d": 0.0009,

            # Coax
            "probe_radius": 0.40,          # mm (a)
            "coax_ba_ratio": 2.3,
            "coax_wall_thickness": 0.20,   # mm
            "coax_port_length": 3.0,       # mm
            "antipad_clearance": 0.10,     # mm

            # Sweep
            "sweep_type": "Interpolating",
            "sweep_step": 0.02,            # GHz
        }

        # ---------- Parâmetros calculados ----------
        self.calculated_params = {
            "num_patches": 4,
            "spacing": 0.0,
            "patch_length": 9.57,
            "patch_width": 9.25,
            "rows": 2,
            "cols": 2,
            "lambda_g": 0.0,
            "substrate_width": 0.0,
            "substrate_length": 0.0
        }

        self.c = 299792458.0

        self.setup_gui()

    # ---------------- GUI ----------------
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        self.window.geometry("1400x900")
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.window, height=64, fg_color=("gray85", "gray20"))
        header.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header.grid_propagate(False)
        ctk.CTkLabel(header, text="Patch Antenna Array Designer",
                     font=ctk.CTkFont(size=26, weight="bold")).pack(pady=12)

        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        for name in ["Parâmetros", "Simulação", "Resultados", "Log"]:
            self.tabview.add(name)
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()

        status = ctk.CTkFrame(self.window, height=36)
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        self.status_label = ctk.CTkLabel(status, text="Pronto.",
                                         font=ctk.CTkFont(weight="bold"))
        self.status_label.pack(pady=6)
        self.process_log_queue()

        # Fecha seguro
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_section(self, parent, title, row, column):
        sec = ctk.CTkFrame(parent, fg_color=("gray92", "gray18"))
        sec.grid(row=row, column=column, sticky="nsew", padx=8, pady=8)
        sec.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(sec, text=title, font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 6))
        ctk.CTkFrame(sec, height=2, fg_color=("gray70", "gray30")).grid(
            row=1, column=0, sticky="ew", padx=10, pady=(0, 6))
        return sec

    def setup_parameters_tab(self):
        tab = self.tabview.tab("Parâmetros")
        main = ctk.CTkScrollableFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.entries = []

        def add_entry(section, label, key, value, row, combo=None, check=False):
            ctk.CTkLabel(section, text=label, font=ctk.CTkFont(weight="bold")).grid(
                row=row, column=0, padx=12, pady=6, sticky="w")
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(section, values=combo, variable=var, width=200)
                widget.grid(row=row, column=1, padx=12, pady=6)
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(section, text="", variable=var)
                widget.grid(row=row, column=1, padx=12, pady=6, sticky="w")
                self.entries.append((key, var))
            else:
                widget = ctk.CTkEntry(section, width=200)
                widget.insert(0, str(value))
                widget.grid(row=row, column=1, padx=12, pady=6)
                self.entries.append((key, widget))
            return row + 1

        # Antena
        sec = self.create_section(main, "Antena", 0, 0)
        r = 2
        r = add_entry(sec, "Frequência central (GHz):", "frequency", self.params["frequency"], r)
        r = add_entry(sec, "Ganho desejado (dBi):", "gain", self.params["gain"], r)
        r = add_entry(sec, "Sweep início (GHz):", "sweep_start", self.params["sweep_start"], r)
        r = add_entry(sec, "Sweep fim (GHz):", "sweep_stop", self.params["sweep_stop"], r)
        r = add_entry(sec, "Espaçamento entre patches:", "spacing_type", self.params["spacing_type"], r,
                      combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        # Substrato
        sec = self.create_section(main, "Substrato", 1, 0)
        r = 2
        r = add_entry(sec, "Material:", "substrate_material", self.params["substrate_material"], r,
                      combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        r = add_entry(sec, "εr:", "er", self.params["er"], r)
        r = add_entry(sec, "tanδ:", "tan_d", self.params["tan_d"], r)
        r = add_entry(sec, "Espessura (mm):", "substrate_thickness", self.params["substrate_thickness"], r)
        r = add_entry(sec, "Metal (mm):", "metal_thickness", self.params["metal_thickness"], r)

        # Coax
        sec = self.create_section(main, "Alimentação coaxial", 2, 0)
        r = 2
        r = add_entry(sec, "Raio interno a (mm):", "probe_radius", self.params["probe_radius"], r)
        r = add_entry(sec, "Razão b/a:", "coax_ba_ratio", self.params["coax_ba_ratio"], r)
        r = add_entry(sec, "Parede do shield (mm):", "coax_wall_thickness", self.params["coax_wall_thickness"], r)
        r = add_entry(sec, "Comprimento do porto Lp (mm):", "coax_port_length", self.params["coax_port_length"], r)
        r = add_entry(sec, "Clear antipad (mm):", "antipad_clearance", self.params["antipad_clearance"], r)

        # Simulação
        sec = self.create_section(main, "Simulação", 3, 0)
        r = 2
        r = add_entry(sec, "CPU Cores:", "cores", self.params["cores"], r)
        r = add_entry(sec, "Salvar projeto:", "save_project", self.save_project, r, check=True)
        r = add_entry(sec, "Sweep:", "sweep_type", self.params["sweep_type"], r,
                      combo=["Discrete", "Interpolating", "Fast"])
        r = add_entry(sec, "Passo Discrete (GHz):", "sweep_step", self.params["sweep_step"], r)

        # Calculados + botões
        sec = self.create_section(main, "Calculados", 4, 0)
        grid = ctk.CTkFrame(sec); grid.grid(row=2, column=0, sticky="nsew", padx=12, pady=8)
        grid.columnconfigure((0,1), weight=1)
        self.patches_label = ctk.CTkLabel(grid, text="Patches: 4", font=ctk.CTkFont(weight="bold")); self.patches_label.grid(row=0, column=0, sticky="w")
        self.rows_cols_label = ctk.CTkLabel(grid, text="Config: 2 x 2", font=ctk.CTkFont(weight="bold")); self.rows_cols_label.grid(row=0, column=1, sticky="w")
        self.spacing_label = ctk.CTkLabel(grid, text="Espaçamento: -- mm", font=ctk.CTkFont(weight="bold")); self.spacing_label.grid(row=1, column=0, sticky="w")
        self.dimensions_label = ctk.CTkLabel(grid, text="Patch: -- x -- mm", font=ctk.CTkFont(weight="bold")); self.dimensions_label.grid(row=1, column=1, sticky="w")
        self.lambda_label = ctk.CTkLabel(grid, text="λg (50Ω): -- mm", font=ctk.CTkFont(weight="bold")); self.lambda_label.grid(row=2, column=0, sticky="w")
        self.substrate_dims_label = ctk.CTkLabel(grid, text="Substrato: -- x -- mm", font=ctk.CTkFont(weight="bold")); self.substrate_dims_label.grid(row=2, column=1, sticky="w")

        btns = ctk.CTkFrame(sec); btns.grid(row=3, column=0, sticky="ew", padx=12, pady=8)
        ctk.CTkButton(btns, text="Calcular", command=self.calculate_parameters, fg_color="#2E8B57").pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Salvar Parâmetros", command=self.save_parameters, fg_color="#4169E1").pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Carregar", command=self.load_parameters, fg_color="#FF8C00").pack(side="left", padx=6)

    def setup_simulation_tab(self):
        tab = self.tabview.tab("Simulação")
        main = ctk.CTkFrame(tab); main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        ctk.CTkLabel(main, text="Controle de simulação", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        row = ctk.CTkFrame(main); row.pack(pady=12)
        self.run_button = ctk.CTkButton(row, text="Executar", command=self.start_simulation_thread, fg_color="#2E8B57", width=160, height=40)
        self.run_button.pack(side="left", padx=6)
        self.stop_button = ctk.CTkButton(row, text="Parar", command=self.stop_simulation_thread, fg_color="#DC143C", state="disabled", width=160, height=40)
        self.stop_button.pack(side="left", padx=6)

        pf = ctk.CTkFrame(main); pf.pack(fill="x", padx=50, pady=8)
        ctk.CTkLabel(pf, text="Progresso:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.progress_bar = ctk.CTkProgressBar(pf, height=18); self.progress_bar.pack(fill="x", pady=6); self.progress_bar.set(0)
        self.sim_status_label = ctk.CTkLabel(main, text="Aguardando…", font=ctk.CTkFont(weight="bold")); self.sim_status_label.pack(pady=8)

    def setup_results_tab(self):
        tab = self.tabview.tab("Resultados")
        main = ctk.CTkFrame(tab); main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1); main.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(main, text="Resultados", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=10)

        g = ctk.CTkFrame(main); g.grid(row=1, column=0, sticky="nsew", padx=10, pady=10); g.grid_columnconfigure(0, weight=1); g.grid_rowconfigure(0, weight=1)
        self.fig = plt.figure(figsize=(10, 9))
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        self.fig.patch.set_facecolor(face)
        self.ax_s11 = self.fig.add_subplot(3, 1, 1)
        self.ax_th = self.fig.add_subplot(3, 1, 2)
        self.ax_ph = self.fig.add_subplot(3, 1, 3)
        for ax in [self.ax_s11, self.ax_th, self.ax_ph]:
            ax.set_facecolor(face)
            ax.grid(True, alpha=0.4)
        self.canvas = FigureCanvasTkAgg(self.fig, master=g); self.canvas.get_tk_widget().pack(fill="both", expand=True)

        ex = ctk.CTkFrame(main); ex.grid(row=2, column=0, pady=8)
        ctk.CTkButton(ex, text="Exportar CSV", command=self.export_csv, fg_color="#6A5ACD").pack(side="left", padx=6)
        ctk.CTkButton(ex, text="Exportar PNG", command=self.export_png, fg_color="#20B2AA").pack(side="left", padx=6)

    def setup_log_tab(self):
        tab = self.tabview.tab("Log")
        main = ctk.CTkFrame(tab); main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1); main.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(main, text="Log", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=10)
        self.log_text = ctk.CTkTextbox(main, width=900, height=500, font=ctk.CTkFont(family="Consolas"))
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log iniciado em " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        btn = ctk.CTkFrame(main); btn.grid(row=2, column=0, pady=8)
        ctk.CTkButton(btn, text="Limpar", command=self.clear_log).pack(side="left", padx=6)
        ctk.CTkButton(btn, text="Salvar", command=self.save_log).pack(side="left", padx=6)

    # ------------- Utilidades de Log -------------
    def log_message(self, message): self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                try:
                    self.log_text.insert("end", msg)
                    self.log_text.see("end")
                except Exception:
                    break
        except queue.Empty:
            pass
        finally:
            try:
                if self.window and self.window.winfo_exists():
                    self.window.after(100, self.process_log_queue)
            except Exception:
                pass

    def clear_log(self): 
        try:
            self.log_text.delete("1.0", "end")
            self.log_message("Log limpo")
        except Exception:
            pass

    def save_log(self):
        try:
            with open("simulation_log.txt", "w", encoding="utf-8") as f:
                f.write(self.log_text.get("1.0", "end"))
            self.log_message("Log salvo em simulation_log.txt")
        except Exception as e:
            self.log_message(f"Erro ao salvar log: {e}")

    def export_csv(self):
        try:
            if self.simulation_data is not None:
                np.savetxt("simulation_results.csv", self.simulation_data, delimiter=",",
                           header="Frequency (GHz), S11 (dB)", comments='')
                self.log_message("CSV exportado.")
            else:
                self.log_message("Sem dados para exportar.")
        except Exception as e:
            self.log_message(f"Erro exportando CSV: {e}")

    def export_png(self):
        try:
            if hasattr(self, 'fig'):
                self.fig.savefig("simulation_results.png", dpi=300, bbox_inches='tight')
                self.log_message("PNG salvo.")
        except Exception as e:
            self.log_message(f"Erro salvando PNG: {e}")

    # ----------- Física / cálculos -----------
    def get_parameters(self):
        self.log_message("Lendo parâmetros…")
        for key, widget in self.entries:
            try:
                if key == "cores":
                    self.params[key] = int(widget.get()) if isinstance(widget, ctk.CTkEntry) else int(self.params[key])
                elif key == "save_project":
                    self.save_project = widget.get()
                elif key in ["substrate_thickness", "metal_thickness", "er", "tan_d",
                             "probe_radius", "coax_ba_ratio", "coax_wall_thickness",
                             "coax_port_length", "antipad_clearance",
                             "sweep_step", "frequency", "gain", "sweep_start", "sweep_stop"]:
                    self.params[key] = float(widget.get()) if isinstance(widget, ctk.CTkEntry) else float(self.params[key])
                elif key in ["spacing_type", "substrate_material", "sweep_type"]:
                    self.params[key] = widget.get()
            except Exception as e:
                msg = f"Valor inválido para {key}: {e}"
                self.status_label.configure(text=msg); self.log_message(msg); return False
        self.log_message("Parâmetros OK.")
        return True

    def _guided_wavelength_ms(self, w_mm: float) -> float:
        """λg (mm) para microstrip com largura w_mm (Hammerstad)."""
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"])  # mm
        w = max(w_mm, 0.01)
        # εeff Hammerstad
        u = w / h
        a = 1 + (1/49.0)*math.log((u**4 + (u/52.0)**2)/(u**4 + 0.432)) + (1/18.7)*math.log(1 + (u/18.1)**3)
        b = 0.564*((er - 0.9)/(er + 3))**0.053
        eeff = (er + 1)/2 + (er - 1)/2*(1 + 10/u)**(-a*b)
        # correção espessura (opcional)
        c0 = self.c
        f = float(self.params["frequency"])*1e9
        lam_g = c0 / (f * math.sqrt(eeff)) * 1000.0  # mm
        return lam_g

    def _microstrip_w(self, Z0_ohm: float) -> float:
        """Hammerstad: devolve W (mm) dado Z0, εr, h(mm)."""
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"])
        Z0 = Z0_ohm
        A = Z0/60.0*math.sqrt((er+1)/2) + (er-1)/(er+1)*(0.23+0.11/er)
        w_h = (8*math.exp(A))/(math.exp(2*A)-2)
        if w_h < 2:
            W = w_h*h
        else:
            B = (377*math.pi)/(2*Z0*math.sqrt(er))
            W = h*(2/math.pi)*(B-1-math.log(2*B-1)+(er-1)/(2*er)*(math.log(B-1)+0.39-0.61/er))
        return max(W, 0.1)

    def calculate_patch_dimensions(self, frequency_ghz: float) -> Tuple[float, float, float]:
        f = frequency_ghz * 1e9
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"]) / 1000.0  # mm->m
        W = self.c / (2 * f) * math.sqrt(2 / (er + 1))
        eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W) ** -0.5
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        L_eff = self.c / (2 * f * math.sqrt(eeff))
        L = L_eff - 2 * dL
        lambda_g = self.c / (f * math.sqrt(eeff))
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)

    def _size_array_from_gain(self):
        G_elem = 8.0
        G_des = float(self.params["gain"])
        N_req = max(1, int(math.ceil(10 ** ((G_des - G_elem) / 10.0))))
        if N_req % 2 == 1: N_req += 1
        rows = max(2, int(round(math.sqrt(N_req))))
        if rows % 2 == 1: rows += 1
        cols = max(2, int(math.ceil(N_req / rows)))
        if cols % 2 == 1: cols += 1
        while rows * cols < N_req:
            if rows <= cols: rows += 2
            else: cols += 2
        return rows, cols, N_req

    def calculate_substrate_size(self):
        L = self.calculated_params["patch_length"]; W = self.calculated_params["patch_width"]
        s = self.calculated_params["spacing"]; r = self.calculated_params["rows"]; c = self.calculated_params["cols"]
        total_w = c * W + (c - 1) * s; total_l = r * L + (r - 1) * s
        margin = max(total_w, total_l) * 0.20
        self.calculated_params["substrate_width"] = total_w + 2 * margin
        self.calculated_params["substrate_length"] = total_l + 2 * margin

    def calculate_parameters(self):
        self.log_message("Calculando parâmetros…")
        if not self.get_parameters(): return
        try:
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params.update({"patch_length": L_mm, "patch_width": W_mm, "lambda_g": lambda_g_mm})
            lambda0_m = self.c / (self.params["frequency"] * 1e9)
            factors = {"lambda/2": 0.5, "lambda": 1.0, "0.7*lambda": 0.7, "0.8*lambda": 0.8, "0.9*lambda": 0.9}
            spacing_mm = factors.get(self.params["spacing_type"], 0.5) * lambda0_m * 1000.0
            self.calculated_params["spacing"] = spacing_mm
            rows, cols, _ = self._size_array_from_gain()
            self.calculated_params.update({"num_patches": rows * cols, "rows": rows, "cols": cols})
            self.calculate_substrate_size()

            self.patches_label.configure(text=f"Patches: {rows*cols}")
            self.rows_cols_label.configure(text=f"Config: {rows} x {cols}")
            self.spacing_label.configure(text=f"Espaçamento: {spacing_mm:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch: {L_mm:.2f} x {W_mm:.2f} mm")
            self.lambda_label.configure(text=f"λg (50Ω): {lambda_g_mm:.2f} mm")
            self.substrate_dims_label.configure(text=f"Substrato: {self.calculated_params['substrate_width']:.2f} x {self.calculated_params['substrate_length']:.2f} mm")
            self.status_label.configure(text="Parâmetros calculados.")
        except Exception as e:
            self.status_label.configure(text=f"Erro no cálculo: {e}")
            self.log_message(f"Erro no cálculo: {e}\n{traceback.format_exc()}")

    # --------- AEDT helpers ---------
    def _ensure_material(self, name: str, er: float, tan_d: float):
        try:
            if not self.hfss.materials.checkifmaterialexists(name):
                self.hfss.materials.add_material(name)
                m = self.hfss.materials.material_keys[name]
                m.permittivity = er
                m.dielectric_loss_tangent = tan_d
                self.log_message(f"Material criado: {name} (er={er}, tanδ={tan_d})")
        except Exception as e:
            self.log_message(f"Material warn '{name}': {e}")

    def _open_project(self):
        self.log_message("Abrindo projeto…")
        if self.desktop is None:
            self.desktop = Desktop(
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"],
                new_desktop=True
            )
            self.log_message("Desktop inicializado.")

        if self.temp_folder is None:
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.project_path = os.path.join(self.temp_folder.name, f"{self.project_title}_{ts}.aedt")

        oDesktop = getattr(self.desktop, "odesktop", None) or getattr(self.desktop, "_odesktop", None)
        oDesktop.NewProject()
        oProject = oDesktop.GetActiveProject()
        if self.design_name not in [d.GetName() for d in oProject.GetDesigns()]:
            oProject.InsertDesign("HFSS", self.design_name, "DrivenModal", "")
        oProject.SetActiveDesign(self.design_name)
        try: oProject.SaveAs(self.project_path, True)
        except Exception: pass

        self.hfss = Hfss(
            project=oProject.GetName(),
            design=self.design_name,
            solution_type="DrivenModal",
            version=self.params["aedt_version"],
            non_graphical=self.params["non_graphical"],
            new_desktop=False
        )
        self.log_message(f"Projeto ativo: {self.project_path}")

    # ---------- utilidades de geometria ----------
    def _safe_name(self, base: str) -> str:
        self._shape_id += 1
        return f"{base}_{self._shape_id:04d}"

    def _rect_h_trace(self, x1: float, x2: float, y: float, w: float, name: str):
        x_min = min(x1, x2); x_max = max(x1, x2)
        origin = [x_min, y - w/2, "h_sub"]
        sizes = [max(x_max - x_min, 1e-6), w]
        return self.hfss.modeler.create_rectangle("XY", origin, sizes, name=self._safe_name(name), material="copper")

    def _rect_v_trace(self, y1: float, y2: float, x: float, w: float, name: str):
        y_min = min(y1, y2); y_max = max(y1, y2)
        origin = [x - w/2, y_min, "h_sub"]
        sizes = [w, max(y_max - y_min, 1e-6)]
        return self.hfss.modeler.create_rectangle("XY", origin, sizes, name=self._safe_name(name), material="copper")

    # ---------- coax ----------
    def _create_coax_feed_lumped(self, ground, substrate, x_feed: float, y_feed: float, name_prefix: str):
        """Apenas o pino interno atravessa o substrato; o externo para no GND."""
        try:
            a_val = float(self.params["probe_radius"])
            b_val = a_val * float(self.params["coax_ba_ratio"])
            wall_val = float(self.params["coax_wall_thickness"])
            Lp_val = float(self.params["coax_port_length"])
            h_sub_val = float(self.params["substrate_thickness"])
            clear_val = float(self.params["antipad_clearance"])
            if b_val - a_val < 0.02: b_val = a_val + 0.02

            # Pino interno: do porto (-Lp) até topo do substrato (h_sub)
            pin = self.hfss.modeler.create_cylinder("Z", [x_feed, y_feed, -Lp_val], a_val,
                                                    h_sub_val + Lp_val + 1e-3,
                                                    name=self._safe_name(f"{name_prefix}_Pin"),
                                                    material="copper")

            # Blindagem externa: fica SOMENTE abaixo do GND (termina em z=0)
            shield_outer = self.hfss.modeler.create_cylinder("Z", [x_feed, y_feed, -Lp_val],
                                                             b_val + wall_val, Lp_val,
                                                             name=self._safe_name(f"{name_prefix}_ShieldOuter"),
                                                             material="copper")
            shield_inner_void = self.hfss.modeler.create_cylinder("Z", [x_feed, y_feed, -Lp_val],
                                                                  b_val, Lp_val,
                                                                  name=self._safe_name(f"{name_prefix}_ShieldInnerVoid"),
                                                                  material="vacuum")
            self.hfss.modeler.subtract(shield_outer, [shield_inner_void], keep_originals=False)

            # Antipad no substrato e furo no GND
            hole_r = b_val + clear_val
            sub_hole = self.hfss.modeler.create_cylinder("Z", [x_feed, y_feed, 0.0], hole_r, h_sub_val,
                                                         name=self._safe_name(f"{name_prefix}_SubHole"), material="vacuum")
            self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=False)
            g_hole = self.hfss.modeler.create_circle("XY", [x_feed, y_feed, 0.0], hole_r,
                                                     name=self._safe_name(f"{name_prefix}_GndHole"), material="vacuum")
            self.hfss.modeler.subtract(ground, [g_hole], keep_originals=False)

            # Anel do porto
            port_ring = self.hfss.modeler.create_circle("XY", [x_feed, y_feed, -Lp_val], b_val,
                                                        name=self._safe_name(f"{name_prefix}_PortRing"), material="vacuum")
            port_hole = self.hfss.modeler.create_circle("XY", [x_feed, y_feed, -Lp_val], a_val,
                                                        name=self._safe_name(f"{name_prefix}_PortHole"), material="vacuum")
            self.hfss.modeler.subtract(port_ring, [port_hole], keep_originals=False)

            # linha de integração (radial)
            eps_line = min(0.1 * (b_val - a_val), 0.05)
            r_start = a_val + eps_line; r_end = b_val - eps_line
            if r_end <= r_start: r_end = a_val + 0.75 * (b_val - a_val)
            p1 = [x_feed + r_start, y_feed, -Lp_val]; p2 = [x_feed + r_end, y_feed, -Lp_val]
            _ = self.hfss.lumped_port(assignment=port_ring.name, integration_line=[p1, p2],
                                      impedance=50.0, name=f"{name_prefix}_Lumped", renormalize=True)

            # pad no topo para soldar na trilha
            top_pad = self.hfss.modeler.create_circle("XY", [x_feed, y_feed, "h_sub"], a_val,
                                                      name=self._safe_name(f"{name_prefix}_TopPad"), material="copper")
            self.log_message(f"Lumped Port {name_prefix}_Lumped criado.")
            return pin, top_pad, shield_outer
        except Exception as e:
            self.log_message(f"Exceção coax '{name_prefix}': {e}\n{traceback.format_exc()}")
            return None, None, None

    # ------------- Simulação -------------
    def start_simulation_thread(self):
        if self.is_simulation_running:
            self.log_message("Simulação já em execução.")
            return
        self.stop_simulation = False
        self.is_simulation_running = True
        threading.Thread(target=self.run_simulation, daemon=True).start()

    def stop_simulation_thread(self):
        self.stop_simulation = True
        self.log_message("Parada solicitada.")

    def run_simulation(self):
        try:
            self.log_message("Iniciando simulação…")
            self.run_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            self.sim_status_label.configure(text="Rodando…")
            self.progress_bar.set(0)

            if not self.get_parameters():
                self.log_message("Parâmetros inválidos."); return
            if self.calculated_params["num_patches"] < 1:
                self.calculate_parameters()

            self._open_project()
            self.progress_bar.set(0.2)
            self.hfss.modeler.model_units = "mm"

            sub_name = self.params["substrate_material"]
            if not self.hfss.materials.checkifmaterialexists(sub_name):
                sub_name = "Custom_Substrate"
                self._ensure_material(sub_name, float(self.params["er"]), float(self.params["tan_d"]))

            L = float(self.calculated_params["patch_length"])
            W = float(self.calculated_params["patch_width"])
            s = float(self.calculated_params["spacing"])
            rows = int(self.calculated_params["rows"])
            cols = int(self.calculated_params["cols"])
            h_sub = float(self.params["substrate_thickness"])
            sub_w = float(self.calculated_params["substrate_width"])
            sub_l = float(self.calculated_params["substrate_length"])

            # Variáveis
            self.hfss["h_sub"] = f"{h_sub}mm"
            self.hfss["t_met"] = f"{self.params['metal_thickness']}mm"
            self.hfss["patchL"] = f"{L}mm"
            self.hfss["patchW"] = f"{W}mm"
            self.hfss["spacing"] = f"{s}mm"
            self.hfss["subW"] = f"{sub_w}mm"
            self.hfss["subL"] = f"{sub_l}mm"
            self.hfss["OLAP"] = "0.10mm"

            # Substrato e ground
            substrate = self.hfss.modeler.create_box(["-subW/2", "-subL/2", 0], ["subW", "subL", "h_sub"],
                                                     name=self._safe_name("Substrate"), material=sub_name)
            ground = self.hfss.modeler.create_rectangle("XY", ["-subW/2", "-subL/2", 0], ["subW", "subL"],
                                                        name=self._safe_name("Ground"), material="copper")

            # Patches (2x2)
            patches = []
            patch_data = []
            cx_left = -(W/2 + s/2)
            cx_right = +(W/2 + s/2)
            cy_bot = -(L/2 + s/2)
            cy_top = +(L/2 + s/2)
            centers = [(cx_left, cy_top), (cx_right, cy_top), (cx_left, cy_bot), (cx_right, cy_bot)]
            for (cx, cy) in centers:
                origin = [cx - W/2, cy - L/2, "h_sub"]
                p = self.hfss.modeler.create_rectangle("XY", origin, ["patchW", "patchL"],
                                                       name=self._safe_name("Patch"), material="copper")
                patches.append(p)
                inner_x = (cx + W/2) if cx < 0 else (cx - W/2)
                patch_data.append({"cx": cx, "cy": cy, "inner_x": inner_x, "obj": p})

            # Porta coaxial + pad
            _, top_pad, _ = self._create_coax_feed_lumped(ground, substrate, 0.0, 0.0, "P0")

            # ------ Rede corporativa paramétrica ------
            # Larguras
            W50 = self._microstrip_w(50.0)
            W70 = self._microstrip_w(70.710678)
            W100 = self._microstrip_w(100.0)
            W200 = self._microstrip_w(200.0)

            # Salva como variáveis (para inspeção no projeto)
            self.hfss["W50"] = f"{W50}mm"; self.hfss["W70"] = f"{W70}mm"; self.hfss["W100"] = f"{W100}mm"; self.hfss["W200"] = f"{W200}mm"

            # λg/4 de 70Ω (depende da largura!)
            Lq70 = self._guided_wavelength_ms(W70)/4.0
            self.hfss["Lq70"] = f"{Lq70}mm"

            # Geometria alvo das colunas (x das colunas verticais)
            xL = -s/2.0; xR = +s/2.0
            y0 = 0.0

            # Tronco: duas seções λ/4 de 70Ω a partir do centro + braços 100Ω até as colunas
            trunk_left_q = self._rect_h_trace(0.0, -Lq70, y0, W70, "Q70_L")
            trunk_right_q = self._rect_h_trace(0.0, +Lq70, y0, W70, "Q70_R")

            # Comprimentos remanescentes 100Ω para alcançar as colunas
            Lh100_L = abs(xL) - Lq70
            Lh100_R = abs(xR) - Lq70
            if Lh100_L < 0: Lh100_L = 0.2  # garante algo positivo
            if Lh100_R < 0: Lh100_R = 0.2
            self.hfss["Lh100_L"] = f"{Lh100_L}mm"
            self.hfss["Lh100_R"] = f"{Lh100_R}mm"
            arm_L = self._rect_h_trace(-Lq70, -Lq70 - Lh100_L, y0, W100, "ARM100_L")
            arm_R = self._rect_h_trace(+Lq70, +Lq70 + Lh100_R, y0, W100, "ARM100_R")

            # λ/4 verticais de 70Ω nas duas colunas
            q_up_L = self._rect_v_trace(y0, y0 + Lq70, xL, W70, "Q70_UP_L")
            q_dn_L = self._rect_v_trace(y0, y0 - Lq70, xL, W70, "Q70_DN_L")
            q_up_R = self._rect_v_trace(y0, y0 + Lq70, xR, W70, "Q70_UP_R")
            q_dn_R = self._rect_v_trace(y0, y0 - Lq70, xR, W70, "Q70_DN_R")

            # Stubs 200Ω até o centro da aresta dos patches (com pequena sobreposição OLAP)
            OLAP = 0.10  # mm
            y_top_edge = cy_top - L/2 + OLAP
            y_bot_edge = cy_bot + L/2 - OLAP

            st_up_L = self._rect_v_trace(y0 + Lq70, y_top_edge, xL, W200, "ST200_UP_L")
            st_dn_L = self._rect_v_trace(y0 - Lq70, y_bot_edge, xL, W200, "ST200_DN_L")
            st_up_R = self._rect_v_trace(y0 + Lq70, y_top_edge, xR, W200, "ST200_UP_R")
            st_dn_R = self._rect_v_trace(y0 - Lq70, y_bot_edge, xR, W200, "ST200_DN_R")

            # Unir metal superior (pad + linhas + patches)
            copper_objs = [top_pad, trunk_left_q, trunk_right_q, arm_L, arm_R,
                           q_up_L, q_dn_L, q_up_R, q_dn_R,
                           st_up_L, st_dn_L, st_up_R, st_dn_R] + patches
            copper_objs = [o for o in copper_objs if o and hasattr(o, "name")]
            try:
                top_metal = self.hfss.modeler.unite(copper_objs)
                top_name = top_metal.name if hasattr(top_metal, "name") else None
            except Exception as e:
                self.log_message(f"Unite aviso: {e}")
                top_name = None

            self.progress_bar.set(0.55)

            # Boundaries
            try:
                names = [ground.name]
                if top_name:
                    names.append(top_name)
                else:
                    names += [o.name for o in copper_objs]
                self.hfss.assign_perfecte_to_sheets(list(dict.fromkeys(names)))
                self.log_message(f"PerfectE em: {names}")
            except Exception as e:
                self.log_message(f"PerfectE aviso: {e}")

            # Região + radiação
            self.log_message("Criando região + radiação")
            lambda0_mm = self.c / (self.params["sweep_start"] * 1e9) * 1000.0
            pad_mm = max(5.0, min(0.25 * float(lambda0_mm), 15.0))
            region = self.hfss.modeler.create_region([pad_mm]*6, is_percentage=False)
            self.hfss.assign_radiation_boundary_to_objects(region)

            # Infinite sphere para far-field
            try:
                rf = self.hfss.odesign.GetModule("RadField")
                rf.InsertInfiniteSphereSetup([
                    "NAME:IS1",
                    "UseCustomRadiationSurface:=", False,
                    "CSDefinition:=", "Theta-Phi",
                    "Polarization:=", "Linear",
                    "ThetaStart:=", "-180deg",
                    "ThetaStop:=", "180deg",
                    "ThetaStep:=", "1deg",
                    "PhiStart:=", "-180deg",
                    "PhiStop:=", "180deg",
                    "PhiStep:=", "1deg",
                    "UseLocalCS:=", False
                ])
            except Exception as e:
                self.log_message(f"InfiniteSphere aviso: {e}")

            self.progress_bar.set(0.65)

            # Setup + sweep
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02
            stype = self.params["sweep_type"]
            try:
                if stype == "Discrete":
                    step = float(self.params["sweep_step"])
                    setup.create_linear_step_sweep(unit="GHz", start_frequency=self.params["sweep_start"],
                                                   stop_frequency=self.params["sweep_stop"], step_size=step, name="Sweep1")
                elif stype == "Fast":
                    setup.create_frequency_sweep(unit="GHz", name="Sweep1",
                                                 start_frequency=self.params["sweep_start"],
                                                 stop_frequency=self.params["sweep_stop"], sweep_type="Fast")
                else:
                    setup.create_frequency_sweep(unit="GHz", name="Sweep1",
                                                 start_frequency=self.params["sweep_start"],
                                                 stop_frequency=self.params["sweep_stop"], sweep_type="Interpolating")
            except Exception as e:
                self.log_message(f"Sweep aviso: {e}")

            # (Opcional) malha local no metal superior
            # if top_name:
            #     try:
            #         lambda_g = max(1e-6, self.calculated_params["lambda_g"])
            #         edge_len = max(lambda_g / 60.0, W / 200.0)
            #         self.hfss.mesh.assign_length_mesh([top_name], maximum_length=f"{edge_len}mm")
            #     except Exception as e:
            #         self.log_message(f"Mesh aviso: {e}")

            # Analisar
            if self.save_project: self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])

            # Pós-processamento nos gráficos
            self.progress_bar.set(0.9)
            self.plot_results()
            self.progress_bar.set(1.0)
            self.sim_status_label.configure(text="Simulação concluída")
            self.log_message("Concluído.")
        except Exception as e:
            self.log_message(f"Erro geral: {e}\n{traceback.format_exc()}")
            self.sim_status_label.configure(text=f"Erro: {e}")
        finally:
            try:
                self.run_button.configure(state="normal")
                self.stop_button.configure(state="disabled")
            except Exception:
                pass
            self.is_simulation_running = False

    # ------------- Plots -------------
    def plot_results(self):
        try:
            self.log_message("Plotando…")
            self.ax_s11.clear(); self.ax_th.clear(); self.ax_ph.clear()

            # S11
            expr = "dB(S(1,1))"
            try:
                rpt = self.hfss.post.reports_by_category.standard(expressions=[expr])
                rpt.context = ["Setup1: Sweep1"]
                sol = rpt.get_solution_data()
            except Exception:
                sol = None

            if sol:
                freqs = np.asarray(sol.primary_sweep_values, dtype=float)
                data = sol.data_real()
                y = np.asarray(data[0] if isinstance(data, (list, tuple)) else data, dtype=float)
                if y.size == freqs.size:
                    self.simulation_data = np.column_stack((freqs, y))
                    self.ax_s11.plot(freqs, y, label="S11", linewidth=2)
                    self.ax_s11.axhline(y=-10, linestyle='--', alpha=0.7, label='-10 dB')
                    cf = float(self.params["frequency"])
                    self.ax_s11.axvline(x=cf, linestyle='--', alpha=0.7)
                    self.ax_s11.set_xlabel("Frequência (GHz)"); self.ax_s11.set_ylabel("dB"); self.ax_s11.set_title("S11"); self.ax_s11.legend(); self.ax_s11.grid(True, alpha=0.5)
            else:
                self.ax_s11.text(0.5, 0.5, "S11 indisponível", transform=self.ax_s11.transAxes, ha="center", va="center")

            # Far-field cortes (se existir)
            try:
                def _cut(cut, fixed):
                    rn = f"Gain_{cut}_{fixed}"
                    exprs = ["dB(GainTotal)"]
                    ctx = "IS1"
                    if cut == "theta":
                        prim = "Theta"; vars_ = {"Freq": f"{self.params['frequency']}GHz", "Phi": f"{fixed}deg", "Theta": "All"}
                    else:
                        prim = "Phi"; vars_ = {"Freq": f"{self.params['frequency']}GHz", "Theta": f"{fixed}deg", "Phi": "All"}
                    rep = self.hfss.post.reports_by_category.far_field(expressions=exprs, context=ctx,
                                                                       primary_sweep_variable=prim,
                                                                       setup="Setup1 : LastAdaptive", variations=vars_, name=rn)
                    sd = rep.get_solution_data()
                    return np.array(sd.primary_sweep_values, dtype=float), np.array(sd.data_real())[0]
                th, gth = _cut("theta", 0.0)
                self.ax_th.plot(th, gth); self.ax_th.set_xlabel("Theta (graus)"); self.ax_th.set_ylabel("Ganho (dB)"); self.ax_th.set_title("Corte Theta, Phi=0"); self.ax_th.grid(True, alpha=0.5)
                ph, gph = _cut("phi", 90.0)
                self.ax_ph.plot(ph, gph); self.ax_ph.set_xlabel("Phi (graus)"); self.ax_ph.set_ylabel("Ganho (dB)"); self.ax_ph.set_title("Corte Phi, Theta=90"); self.ax_ph.grid(True, alpha=0.5)
            except Exception as e:
                self.log_message(f"Far-field indisponível: {e}")
                self.ax_th.text(0.5, 0.5, "FF indisponível", transform=self.ax_th.transAxes, ha="center", va="center")
                self.ax_ph.text(0.5, 0.5, "FF indisponível", transform=self.ax_ph.transAxes, ha="center", va="center")

            self.fig.tight_layout(); self.canvas.draw()
            self.log_message("Plot OK.")
        except Exception as e:
            self.log_message(f"Erro nos gráficos: {e}\n{traceback.format_exc()}")

    # ------------- Cleanup -------------
    def cleanup(self):
        try:
            if self.hfss:
                try:
                    if self.save_project: self.hfss.save_project()
                    else: self.hfss.close_project(save=False)
                except Exception as e: self.log_message(f"Erro ao fechar projeto: {e}")
            if self.desktop:
                try: self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception as e: self.log_message(f"Erro ao liberar desktop: {e}")
            if self.temp_folder and not self.save_project:
                try: self.temp_folder.cleanup()
                except Exception as e: self.log_message(f"Erro limpando temporários: {e}")
        except Exception as e:
            self.log_message(f"Cleanup erro: {e}")

    def on_closing(self):
        self.log_message("Fechando…")
        try:
            self.window.after_cancel(self.process_log_queue)
        except Exception:
            pass
        self.cleanup()
        try:
            self.window.destroy()
        except Exception:
            pass

    # ------------- Persistência -------------
    def save_parameters(self):
        try:
            all_params = {**self.params, **self.calculated_params}
            with open("antenna_parameters.json", "w") as f: json.dump(all_params, f, indent=4)
            self.log_message("Parâmetros salvos.")
        except Exception as e: self.log_message(f"Erro ao salvar parâmetros: {e}")

    def load_parameters(self):
        try:
            with open("antenna_parameters.json", "r") as f: all_params = json.load(f)
            for k in self.params:
                if k in all_params: self.params[k] = all_params[k]
            for k in self.calculated_params:
                if k in all_params: self.calculated_params[k] = all_params[k]
            self.update_interface_from_params()
            self.log_message("Parâmetros carregados.")
        except Exception as e:
            self.log_message(f"Erro ao carregar parâmetros: {e}")

    def update_interface_from_params(self):
        try:
            for key, widget in self.entries:
                if key in self.params:
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end"); widget.insert(0, str(self.params[key]))
                    elif isinstance(widget, ctk.StringVar):
                        widget.set(self.params[key])
                    elif isinstance(widget, ctk.BooleanVar):
                        widget.set(self.params[key])
            self.patches_label.configure(text=f"Patches: {self.calculated_params['num_patches']}")
            self.rows_cols_label.configure(text=f"Config: {self.calculated_params['rows']} x {self.calculated_params['cols']}")
            self.spacing_label.configure(text=f"Espaçamento: {self.calculated_params['spacing']:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch: {self.calculated_params['patch_length']:.2f} x {self.calculated_params['patch_width']:.2f} mm")
            self.lambda_label.configure(text=f"λg (50Ω): {self.calculated_params['lambda_g']:.2f} mm")
            self.substrate_dims_label.configure(text=f"Substrato: {self.calculated_params['substrate_width']:.2f} x {self.calculated_params['substrate_length']:.2f} mm")
        except Exception as e:
            self.log_message(f"Erro atualizando GUI: {e}")

    def run(self):
        try:
            self.window.mainloop()
        finally:
            self.cleanup()


if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()
