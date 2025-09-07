# -*- coding: utf-8 -*-
import os
import re
import tempfile
from datetime import datetime
import math
import json
import traceback
import queue
import threading
from typing import Tuple, List, Dict, Optional

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk

import ansys.aedt.core
from ansys.aedt.core import Desktop, Hfss

# ---------- Aparência ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ModernPatchAntennaDesigner:
    def __init__(self):
        self.hfss: Optional[Hfss] = None
        self.desktop: Optional[Desktop] = None
        self.temp_folder = None
        self.project_path = ""
        self.project_display_name = "patch_array"
        self.design_base_name = "patch_array"
        self.log_queue = queue.Queue()
        self.is_simulation_running = False
        self.save_project = False
        self.stop_simulation = False

        # excitações
        self.saved_excitations: List[str] = []
        self.port_vars: Dict[str, Tuple[str, str]] = {}

        # última análise S11
        self.last_analysis = {
            "port": None,
            "f_res": None,
            "s11_min_db": None,
            "R": None,
            "X": None,
            "scale": None
        }

        # -------- Parâmetros do usuário --------
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
            "feed_position": "inset",      # edge|inset
            "feed_rel_x": 0.485,
            "probe_radius": 0.40,          # mm (a)
            "coax_ba_ratio": 2.3,
            "coax_wall_thickness": 0.20,   # mm
            "coax_port_length": 3.0,       # mm
            "antipad_clearance": 0.10,     # mm

            "sweep_type": "Interpolating", # "Discrete" | "Interpolating" | "Fast"
            "sweep_step": 0.02,            # GHz (só para Discrete)

            "plot_results": True           # habilitei por padrão para ver S11 + marcação
        }

        # -------- Parâmetros calculados --------
        self.calculated_params = {
            "num_patches": 4,
            "spacing": 0.0,
            "patch_length": 9.57,
            "patch_width": 9.25,
            "rows": 2,
            "cols": 2,
            "lambda_g": 0.0,
            "feed_offset": 2.0,
            "substrate_width": 0.0,
            "substrate_length": 0.0
        }

        self.c = 299792458.0
        self._build_ui()

    # ---------------- GUI ----------------
    def _build_ui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        try:
            self.window.state('zoomed')
        except Exception:
            try:
                self.window.attributes('-zoomed', True)
            except Exception:
                self.window.geometry("1400x900")
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.window, height=64, fg_color=("gray85", "gray20"))
        header.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header.grid_propagate(False)
        ctk.CTkLabel(header, text="Patch Antenna Array Designer",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=("gray10", "gray90")).pack(pady=14)

        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        for name in ["Design Parameters", "Simulation", "Excitations", "Results", "Log"]:
            self.tabview.add(name)
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        self._tab_parameters()
        self._tab_simulation()
        self._tab_excitations()
        self._tab_results()
        self._tab_log()

        status = ctk.CTkFrame(self.window, height=40)
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        self.status_label = ctk.CTkLabel(status, text="Ready to calculate parameters",
                                         font=ctk.CTkFont(weight="bold"))
        self.status_label.pack(pady=8)

        self._drain_log_queue()

    def _section(self, parent, title, row, column, padx=10, pady=10):
        section = ctk.CTkFrame(parent, fg_color=("gray92", "gray18"))
        section.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
        section.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(section, text=title,
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=("gray20", "gray80")).grid(row=0, column=0, sticky="w", padx=15, pady=(10, 6))
        ctk.CTkFrame(section, height=2, fg_color=("gray70", "gray30")).grid(row=1, column=0, sticky="ew", padx=10)
        return section

    def _tab_parameters(self):
        tab = self.tabview.tab("Design Parameters")
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkScrollableFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main.grid_columnconfigure(0, weight=1)

        sec_ant = self._section(main, "Antenna Parameters", 0, 0)
        entries = []
        r = 2

        def add(section, label, key, value, row, combo=None, check=False):
            ctk.CTkLabel(section, text=label, font=ctk.CTkFont(weight="bold")
                         ).grid(row=row, column=0, padx=15, pady=6, sticky="w")
            if combo:
                var = ctk.StringVar(value=value)
                w = ctk.CTkComboBox(section, values=combo, variable=var, width=220)
                w.grid(row=row, column=1, padx=15, pady=6)
                entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                w = ctk.CTkCheckBox(section, text="", variable=var)
                w.grid(row=row, column=1, padx=15, pady=6, sticky="w")
                entries.append((key, var))
            else:
                w = ctk.CTkEntry(section, width=220)
                w.insert(0, str(value))
                w.grid(row=row, column=1, padx=15, pady=6)
                entries.append((key, w))
            return row + 1

        r = add(sec_ant, "Central Frequency (GHz):", "frequency", self.params["frequency"], r)
        r = add(sec_ant, "Desired Gain (dBi):", "gain", self.params["gain"], r)
        r = add(sec_ant, "Sweep Start (GHz):", "sweep_start", self.params["sweep_start"], r)
        r = add(sec_ant, "Sweep Stop (GHz):", "sweep_stop", self.params["sweep_stop"], r)
        r = add(sec_ant, "Patch Spacing:", "spacing_type", self.params["spacing_type"], r,
                combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        sec_sub = self._section(main, "Substrate Parameters", 1, 0)
        r = 2
        r = add(sec_sub, "Substrate Material:", "substrate_material",
                self.params["substrate_material"], r,
                combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        r = add(sec_sub, "Relative Permittivity (εr):", "er", self.params["er"], r)
        r = add(sec_sub, "Loss Tangent (tan δ):", "tan_d", self.params["tan_d"], r)
        r = add(sec_sub, "Substrate Thickness (mm):", "substrate_thickness", self.params["substrate_thickness"], r)
        r = add(sec_sub, "Metal Thickness (mm):", "metal_thickness", self.params["metal_thickness"], r)

        sec_coax = self._section(main, "Coaxial Feed Parameters", 2, 0)
        r = 2
        r = add(sec_coax, "Feed position type:", "feed_position", self.params["feed_position"], r,
                combo=["inset", "edge"])
        r = add(sec_coax, "Feed relative X (0..1):", "feed_rel_x", self.params["feed_rel_x"], r)
        r = add(sec_coax, "Inner radius a (mm):", "probe_radius", self.params["probe_radius"], r)
        r = add(sec_coax, "b/a ratio:", "coax_ba_ratio", self.params["coax_ba_ratio"], r)
        r = add(sec_coax, "Shield wall (mm):", "coax_wall_thickness", self.params["coax_wall_thickness"], r)
        r = add(sec_coax, "Port length below GND Lp (mm):", "coax_port_length", self.params["coax_port_length"], r)
        r = add(sec_coax, "Anti-pad clearance (mm):", "antipad_clearance", self.params["antipad_clearance"], r)

        sec_sim = self._section(main, "Simulation Settings", 3, 0)
        r = 2
        r = add(sec_sim, "CPU Cores:", "cores", self.params["cores"], r)
        r = add(sec_sim, "Show HFSS Interface:", "show_gui", not self.params["non_graphical"], r, check=True)
        r = add(sec_sim, "Save Project:", "save_project", self.save_project, r, check=True)
        r = add(sec_sim, "Sweep Type:", "sweep_type", self.params["sweep_type"], r,
                combo=["Discrete", "Interpolating", "Fast"])
        r = add(sec_sim, "Discrete Step (GHz):", "sweep_step", self.params["sweep_step"], r)
        r = add(sec_sim, "Plot Results After Solve:", "plot_results", self.params["plot_results"], r, check=True)

        self.entries = entries

        sec_calc = self._section(main, "Calculated Parameters", 4, 0)
        grid = ctk.CTkFrame(sec_calc)
        grid.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        grid.columnconfigure((0, 1), weight=1)

        self.patches_label = ctk.CTkLabel(grid, text="Number of Patches: 4", font=ctk.CTkFont(weight="bold"))
        self.patches_label.grid(row=0, column=0, sticky="w", pady=4)
        self.rows_cols_label = ctk.CTkLabel(grid, text="Configuration: 2 x 2", font=ctk.CTkFont(weight="bold"))
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=4)
        self.spacing_label = ctk.CTkLabel(grid, text="Spacing: -- mm", font=ctk.CTkFont(weight="bold"))
        self.spacing_label.grid(row=1, column=0, sticky="w", pady=4)
        self.dimensions_label = ctk.CTkLabel(grid, text="Patch Dimensions: -- x -- mm",
                                             font=ctk.CTkFont(weight="bold"))
        self.dimensions_label.grid(row=1, column=1, sticky="w", pady=4)
        self.lambda_label = ctk.CTkLabel(grid, text="Guided Wavelength: -- mm", font=ctk.CTkFont(weight="bold"))
        self.lambda_label.grid(row=2, column=0, sticky="w", pady=4)
        self.feed_offset_label = ctk.CTkLabel(grid, text="Feed Offset (y): -- mm", font=ctk.CTkFont(weight="bold"))
        self.feed_offset_label.grid(row=2, column=1, sticky="w", pady=4)
        self.substrate_dims_label = ctk.CTkLabel(grid, text="Substrate Dimensions: -- x -- mm",
                                                 font=ctk.CTkFont(weight="bold"))
        self.substrate_dims_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=4)

        btns = ctk.CTkFrame(sec_calc)
        btns.grid(row=3, column=0, sticky="ew", padx=15, pady=12)
        ctk.CTkButton(btns, text="Calculate Parameters", command=self.calculate_parameters,
                      fg_color="#2E8B57", hover_color="#3CB371", width=180).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="Save Parameters", command=self.save_parameters,
                      fg_color="#4169E1", hover_color="#6495ED", width=140).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="Load Parameters", command=self.load_parameters,
                      fg_color="#FF8C00", hover_color="#FFA500", width=140).pack(side="left", padx=8)

    def _tab_simulation(self):
        tab = self.tabview.tab("Simulation")
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(main, text="Simulation Control", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)

        row = ctk.CTkFrame(main)
        row.pack(pady=14)
        self.run_button = ctk.CTkButton(row, text="Run Simulation", command=self.start_simulation_thread,
                                        fg_color="#2E8B57", hover_color="#3CB371", height=40, width=160)
        self.run_button.pack(side="left", padx=8)
        self.stop_button = ctk.CTkButton(row, text="Stop Simulation", command=self.stop_simulation_thread,
                                         fg_color="#DC143C", hover_color="#FF4500",
                                         state="disabled", height=40, width=160)
        self.stop_button.pack(side="left", padx=8)

        bar = ctk.CTkFrame(main)
        bar.pack(fill="x", padx=50, pady=8)
        ctk.CTkLabel(bar, text="Simulation Progress:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.progress_bar = ctk.CTkProgressBar(bar, height=18)
        self.progress_bar.pack(fill="x", pady=6)
        self.progress_bar.set(0)

        self.sim_status_label = ctk.CTkLabel(main, text="Simulation not started",
                                             font=ctk.CTkFont(weight="bold"))
        self.sim_status_label.pack(pady=8)

        note = ctk.CTkFrame(main, fg_color=("gray90", "gray15"))
        note.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(note, text="Obs.: pós-solve criamos variáveis p_i/ph_i e esfera infinita; "
                                 "agora também analisamos S11 e sugerimos correção geométrica.",
                     font=ctk.CTkFont(size=12, slant="italic"),
                     text_color=("gray40", "gray60")).pack(padx=10, pady=10)

    def _tab_excitations(self):
        tab = self.tabview.tab("Excitations")
        tab.grid_rowconfigure(2, weight=1)

        top = self._section(tab, "Ports & Post-Processing Variables", 0, 0)
        self.ports_info_label = ctk.CTkLabel(top, text="Ports: 0", font=ctk.CTkFont(weight="bold"))
        self.ports_info_label.grid(row=2, column=0, padx=15, pady=6, sticky="w")

        btns = ctk.CTkFrame(top)
        btns.grid(row=3, column=0, sticky="w", padx=15, pady=8)
        ctk.CTkButton(btns, text="Refresh Ports", command=self.refresh_ports_ui,
                      fg_color="#1E88E5", hover_color="#42A5F5").pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Create/Sync Vars", command=self.create_sync_post_vars,
                      fg_color="#2E8B57", hover_color="#3CB371").pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Apply to HFSS", command=self.apply_sources_from_vars,
                      fg_color="#6A1B9A", hover_color="#8E24AA").pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Equal Power (1W)", command=self.set_equal_power_1w,
                      fg_color="#546E7A", hover_color="#78909C").pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Zero All Phases", command=self.zero_all_phases,
                      fg_color="#546E7A", hover_color="#78909C").pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Update Var Values", command=self.update_var_values_in_aedt,
                      fg_color="#EF6C00", hover_color="#FB8C00").pack(side="left", padx=6)

        # tabela
        self.ex_table = ctk.CTkScrollableFrame(tab, fg_color=("gray93", "gray17"))
        self.ex_table.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        self.ex_table.grid_columnconfigure(0, weight=1)

        hdr = ctk.CTkFrame(self.ex_table, fg_color=("gray85", "gray25"))
        hdr.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        for j, t in enumerate(["Port", "Power Var", "Power Value", "Phase Var", "Phase Value"]):
            ctk.CTkLabel(hdr, text=t, font=ctk.CTkFont(weight="bold")).grid(row=0, column=j, padx=10, pady=6)
            hdr.grid_columnconfigure(j, weight=1)

        self.ex_rows: Dict[str, Dict[str, ctk.CTkEntry]] = {}

    def _tab_results(self):
        tab = self.tabview.tab("Results")
        tab.grid_rowconfigure(1, weight=1)

        top = self._section(tab, "S11 Analysis & Resonance Tuning", 0, 0)
        row = ctk.CTkFrame(top)
        row.grid(row=2, column=0, sticky="w", padx=15, pady=8)
        ctk.CTkButton(row, text="Analyze S11 (auto)", command=self.analyze_and_mark_s11,
                      fg_color="#1976D2", hover_color="#42A5F5").pack(side="left", padx=6)
        ctk.CTkButton(row, text="Apply Resonance Correction", command=self.apply_resonance_correction,
                      fg_color="#2E7D32", hover_color="#43A047").pack(side="left", padx=6)

        self.analysis_label = ctk.CTkLabel(top, text="—", font=ctk.CTkFont(weight="bold"))
        self.analysis_label.grid(row=3, column=0, sticky="w", padx=15, pady=6)

        main = ctk.CTkFrame(tab)
        main.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)

        self.fig = plt.figure(figsize=(9, 6))
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        self.fig.patch.set_facecolor(face)
        self.ax_s11 = self.fig.add_subplot(1, 1, 1)
        self.ax_s11.set_facecolor(face)
        if ctk.get_appearance_mode() == "Dark":
            self.ax_s11.tick_params(colors='white')
            self.ax_s11.xaxis.label.set_color('white')
            self.ax_s11.yaxis.label.set_color('white')
            self.ax_s11.title.set_color('white')
            for s in ['bottom', 'top', 'right', 'left']:
                self.ax_s11.spines[s].set_color('white')
            self.ax_s11.grid(color='gray', alpha=0.5)

        self.canvas = FigureCanvasTkAgg(self.fig, master=main)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def _tab_log(self):
        tab = self.tabview.tab("Log")
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(main, text="Simulation Log", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0,
                                                                                                  pady=10)
        self.log_text = ctk.CTkTextbox(main, width=900, height=500, font=ctk.CTkFont(family="Consolas"))
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        btn = ctk.CTkFrame(main)
        btn.grid(row=2, column=0, pady=8)
        ctk.CTkButton(btn, text="Clear Log", command=self.clear_log).pack(side="left", padx=8)
        ctk.CTkButton(btn, text="Save Log", command=self.save_log).pack(side="left", padx=8)

    # ------------- Log utils -------------
    def log_message(self, message):
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")

    def _drain_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert("end", msg)
                self.log_text.see("end")
        except queue.Empty:
            pass
        finally:
            if self.window.winfo_exists():
                self.window.after(100, self._drain_log_queue)

    def clear_log(self):
        self.log_text.delete("1.0", "end")
        self.log_message("Log cleared")

    def save_log(self):
        try:
            with open("simulation_log.txt", "w", encoding="utf-8") as f:
                f.write(self.log_text.get("1.0", "end"))
            self.log_message("Log saved to simulation_log.txt")
        except Exception as e:
            self.log_message(f"Error saving log: {e}")

    # ----------- Física / cálculos -----------
    def get_parameters(self):
        self.log_message("Getting parameters from interface")
        for key, widget in self.entries:
            try:
                if key == "cores":
                    self.params[key] = int(widget.get()) if isinstance(widget, ctk.CTkEntry) else int(self.params[key])
                elif key == "show_gui":
                    self.params["non_graphical"] = not widget.get()
                elif key == "save_project":
                    self.save_project = widget.get()
                elif key in ["substrate_thickness", "metal_thickness", "er", "tan_d",
                             "probe_radius", "coax_ba_ratio", "coax_wall_thickness",
                             "coax_port_length", "antipad_clearance", "feed_rel_x",
                             "sweep_step"]:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
                elif key in ["spacing_type", "substrate_material", "feed_position", "sweep_type"]:
                    self.params[key] = widget.get()
                elif key == "plot_results":
                    self.params[key] = widget.get()
                else:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
            except Exception as e:
                msg = f"Invalid value for {key}: {e}"
                self.status_label.configure(text=msg)
                self.log_message(msg)
                return False
        self.log_message("All parameters retrieved successfully")
        return True

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
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)  # mm

    def _size_array_from_gain(self):
        G_elem = 8.0  # dBi
        G_des = float(self.params["gain"])
        N_req = max(1, int(math.ceil(10 ** ((G_des - G_elem) / 10.0))))
        if N_req % 2 == 1:
            N_req += 1
        rows = max(2, int(round(math.sqrt(N_req))))
        if rows % 2 == 1:
            rows += 1
        cols = max(2, int(math.ceil(N_req / rows)))
        if cols % 2 == 1:
            cols += 1
        while rows * cols < N_req:
            if rows <= cols:
                rows += 2
            else:
                cols += 2
        return rows, cols, N_req

    def calculate_substrate_size(self):
        L = self.calculated_params["patch_length"]
        W = self.calculated_params["patch_width"]
        s = self.calculated_params["spacing"]
        r = self.calculated_params["rows"]
        c = self.calculated_params["cols"]
        total_w = c * W + (c - 1) * s
        total_l = r * L + (r - 1) * s
        margin = max(total_w, total_l) * 0.20
        self.calculated_params["substrate_width"] = total_w + 2 * margin
        self.calculated_params["substrate_length"] = total_l + 2 * margin
        self.log_message(f"Substrate size calculated: {self.calculated_params['substrate_width']:.2f} x "
                         f"{self.calculated_params['substrate_length']:.2f} mm")

    def calculate_parameters(self):
        self.log_message("Starting parameter calculation")
        if not self.get_parameters():
            self.log_message("Parameter calculation failed due to invalid input")
            return
        try:
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params.update({"patch_length": L_mm, "patch_width": W_mm, "lambda_g": lambda_g_mm})
            lambda0_m = self.c / (self.params["frequency"] * 1e9)
            factors = {"lambda/2": 0.5, "lambda": 1.0, "0.7*lambda": 0.7, "0.8*lambda": 0.8, "0.9*lambda": 0.9}
            spacing_mm = factors.get(self.params["spacing_type"], 0.5) * lambda0_m * 1000.0
            self.calculated_params["spacing"] = spacing_mm
            rows, cols, N_req = self._size_array_from_gain()
            self.calculated_params.update({"num_patches": rows * cols, "rows": rows, "cols": cols})
            self.log_message(f"Array sizing -> target gain {self.params['gain']} dBi, N_req≈{N_req}, "
                             f"layout {rows}x{cols} (= {rows*cols} patches)")
            self.calculated_params["feed_offset"] = 0.30 * L_mm
            self.calculate_substrate_size()
            self.patches_label.configure(text=f"Number of Patches: {rows*cols}")
            self.rows_cols_label.configure(text=f"Configuration: {rows} x {cols}")
            self.spacing_label.configure(text=f"Spacing: {spacing_mm:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch Dimensions: {L_mm:.2f} x {W_mm:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {lambda_g_mm:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset (y): {self.calculated_params['feed_offset']:.2f} mm")
            self.substrate_dims_label.configure(
                text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x "
                     f"{self.calculated_params['substrate_length']:.2f} mm")
            self.status_label.configure(text="Parameters calculated successfully")
            self.log_message("Parameters calculated successfully")
        except Exception as e:
            self.status_label.configure(text=f"Error in calculation: {e}")
            self.log_message(f"Error in calculation: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # --------- AEDT helpers ---------
    def _ensure_material(self, name: str, er: float, tan_d: float):
        try:
            if not self.hfss.materials.checkifmaterialexists(name):
                self.hfss.materials.add_material(name)
                m = self.hfss.materials.material_keys[name]
                m.permittivity = er
                m.dielectric_loss_tangent = tan_d
                self.log_message(f"Created material: {name} (er={er}, tanδ={tan_d})")
        except Exception as e:
            self.log_message(f"Material management warning for '{name}': {e}")

    def _open_or_create_project(self):
        if self.desktop is None:
            self.desktop = Desktop(version=self.params["aedt_version"],
                                   non_graphical=self.params["non_graphical"],
                                   new_desktop=True)
        od = getattr(self.desktop, "_odesktop", None) or getattr(self.desktop, "odesktop", None)
        open_names, open_objs = [], []
        if od:
            try:
                for p in od.GetProjects():
                    try:
                        open_names.append(p.GetName())
                        open_objs.append(p)
                    except Exception:
                        pass
            except Exception:
                pass
        if self.project_display_name in open_names:
            idx = open_names.index(self.project_display_name)
            proj_obj = open_objs[idx]
            new_design = self.design_base_name
            try:
                tmp = Hfss(project=proj_obj, non_graphical=self.params["non_graphical"],
                           version=self.params["aedt_version"])
                existing = list(tmp.project.design_list)
                k = 1
                while new_design in existing:
                    k += 1
                    new_design = f"{self.design_base_name}_{k}"
                try:
                    tmp.close_desktop()
                except Exception:
                    pass
            except Exception:
                new_design = f"{self.design_base_name}_{datetime.now().strftime('%H%M%S')}"
            self.hfss = Hfss(project=proj_obj, design=new_design, solution_type="DrivenModal",
                             version=self.params["aedt_version"], non_graphical=self.params["non_graphical"])
            try:
                self.project_path = proj_obj.GetPath()
            except Exception:
                self.project_path = ""
            self.log_message(f"Using existing project '{self.project_display_name}', created design '{new_design}'")
            return
        if self.temp_folder is None:
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.project_path = os.path.join(self.temp_folder.name, f"{self.project_display_name}_{ts}.aedt")
        self.hfss = Hfss(project=self.project_path, design=self.design_base_name, solution_type="DrivenModal",
                         version=self.params["aedt_version"], non_graphical=self.params["non_graphical"])
        self.log_message(f"Created new project: {self.project_path} (design '{self.design_base_name}')")

    def _set_design_variables(self, L, W, spacing, rows, cols, h_sub, sub_w, sub_l):
        a = float(self.params["probe_radius"])
        ba = float(self.params["coax_ba_ratio"])
        b = a * ba
        wall = float(self.params["coax_wall_thickness"])
        Lp = float(self.params["coax_port_length"])
        clear = float(self.params["antipad_clearance"])

        self.hfss["f0"] = f"{self.params['frequency']}GHz"
        self.hfss["h_sub"] = f"{h_sub}mm"
        self.hfss["t_met"] = f"{self.params['metal_thickness']}mm"
        self.hfss["patchL"] = f"{L}mm"
        self.hfss["patchW"] = f"{W}mm"
        self.hfss["spacing"] = f"{spacing}mm"
        self.hfss["rows"] = str(rows)
        self.hfss["cols"] = str(cols)
        self.hfss["subW"] = f"{sub_w}mm"
        self.hfss["subL"] = f"{sub_l}mm"
        self.hfss["a"] = f"{a}mm"
        self.hfss["b"] = f"{b}mm"
        self.hfss["wall"] = f"{wall}mm"
        self.hfss["Lp"] = f"{Lp}mm"
        self.hfss["clear"] = f"{clear}mm"
        self.hfss["eps"] = "0.001mm"
        self.hfss["padAir"] = f"{max(spacing, W, L)/2 + Lp + 2.0}mm"
        self.log_message(f"Air coax set: a={a:.3f} mm, b={b:.3f} mm (b/a={ba:.3f}≈2.3 → ~50 Ω)")
        return a, b, wall, Lp, clear

    def _create_coax_feed_lumped(self, ground, substrate, x_feed: float, y_feed: float, name_prefix: str):
        try:
            a_val = float(self.params["probe_radius"])
            b_val = a_val * float(self.params["coax_ba_ratio"])
            wall_val = float(self.params["coax_wall_thickness"])
            Lp_val = float(self.params["coax_port_length"])
            h_sub_val = float(self.params["substrate_thickness"])
            clear_val = float(self.params["antipad_clearance"])
            if b_val - a_val < 0.02:
                b_val = a_val + 0.02

            pin = self.hfss.modeler.create_cylinder(
                orientation="Z", origin=[x_feed, y_feed, -Lp_val],
                radius=a_val, height=h_sub_val + Lp_val + 0.001,
                name=f"{name_prefix}_Pin", material="copper")

            shield_outer = self.hfss.modeler.create_cylinder(
                orientation="Z", origin=[x_feed, y_feed, -Lp_val],
                radius=b_val + wall_val, height=Lp_val,
                name=f"{name_prefix}_ShieldOuter", material="copper")
            shield_inner_void = self.hfss.modeler.create_cylinder(
                orientation="Z", origin=[x_feed, y_feed, -Lp_val],
                radius=b_val, height=Lp_val,
                name=f"{name_prefix}_ShieldInnerVoid", material="vacuum")
            self.hfss.modeler.subtract(shield_outer, [shield_inner_void], keep_originals=False)

            hole_r = b_val + clear_val
            sub_hole = self.hfss.modeler.create_cylinder(
                orientation="Z", origin=[x_feed, y_feed, 0.0],
                radius=hole_r, height=h_sub_val, name=f"{name_prefix}_SubHole", material="vacuum")
            self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=False)
            g_hole = self.hfss.modeler.create_circle(
                orientation="XY", origin=[x_feed, y_feed, 0.0], radius=hole_r,
                name=f"{name_prefix}_GndHole", material="vacuum")
            self.hfss.modeler.subtract(ground, [g_hole], keep_originals=False)

            port_ring = self.hfss.modeler.create_circle(
                orientation="XY", origin=[x_feed, y_feed, -Lp_val],
                radius=b_val, name=f"{name_prefix}_PortRing", material="vacuum")
            port_hole = self.hfss.modeler.create_circle(
                orientation="XY", origin=[x_feed, y_feed, -Lp_val],
                radius=a_val, name=f"{name_prefix}_PortHole", material="vacuum")
            self.hfss.modeler.subtract(port_ring, [port_hole], keep_originals=False)

            eps_line = min(0.1 * (b_val - a_val), 0.05)
            r_start = a_val + eps_line
            r_end = b_val - eps_line
            if r_end <= r_start:
                r_end = a_val + 0.75 * (b_val - a_val)
            p1 = [x_feed + r_start, y_feed, -Lp_val]
            p2 = [x_feed + r_end,   y_feed, -Lp_val]

            self.hfss.lumped_port(
                assignment=port_ring.name,
                integration_line=[p1, p2],
                impedance=50.0,
                name=f"{name_prefix}_Lumped",
                renormalize=True
            )
            portname = f"{name_prefix}_Lumped:1"
            self.saved_excitations.append(portname)
            self.log_message(f"Lumped Port '{portname}' created (integration line).")
            return pin, None, shield_outer
        except Exception as e:
            self.log_message(f"Exception in coax creation '{name_prefix}': {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            return None, None, None

    # ---------- Pós-solve helpers ----------
    def _change_property(self, tab_payload):
        try:
            self.hfss.odesign.ChangeProperty(tab_payload)
            return True
        except Exception as e:
            self.log_message(f"ChangeProperty failed: {e}")
            return False

    def _ensure_post_var(self, name: str, value: str):
        create = [
            "NAME:AllTabs",
            [
                "NAME:LocalVariableTab",
                ["NAME:PropServers", "LocalVariables"],
                ["NAME:NewProps",
                 [f"NAME:{name}", "PropType:=", "PostProcessingVariableProp", "UserDef:=", True, "Value:=", value]]
            ]
        ]
        if self._change_property(create):
            self.log_message(f"Post var '{name}' = {value} created.")
            return True
        modify = [
            "NAME:AllTabs",
            [
                "NAME:LocalVariableTab",
                ["NAME:PropServers", "LocalVariables"],
                ["NAME:ChangedProps", [f"NAME:{name}", "Value:=", value]]
            ]
        ]
        if self._change_property(modify):
            self.log_message(f"Post var '{name}' updated to {value}.")
            return True
        return False

    @staticmethod
    def _port_index_from_name(port: str) -> Optional[int]:
        m = re.search(r"P(\d+)_Lumped", port, re.IGNORECASE)
        return int(m.group(1)) if m else None

    def _ensure_vars_for_ports(self, excitations: List[str]):
        mapping: Dict[str, Tuple[str, str]] = {}
        if not excitations:
            return mapping
        def sort_key(p):  # ordena por índice Pn
            idx = self._port_index_from_name(p)
            return idx if idx is not None else 9999
        exs = sorted(excitations, key=sort_key)
        for k, ex in enumerate(exs, start=1):
            idx = self._port_index_from_name(ex) or k
            pvar_simple = f"p{idx}"
            phvar_simple = f"ph{idx}"
            pvar_alt = f"Pow_P{idx}"
            phvar_alt = f"Phi_P{idx}"
            self._ensure_post_var(pvar_simple, "1W")
            self._ensure_post_var(phvar_simple, "0deg")
            self._ensure_post_var(pvar_alt, "1W")
            self._ensure_post_var(phvar_alt, "0deg")
            mapping[ex] = (pvar_simple, phvar_simple)
        self.port_vars = mapping
        return mapping

    def _apply_edit_sources(self, mapping: Dict[str, Tuple[str, str]]):
        if not mapping:
            self.log_message("Apply skipped: no port/variable mapping.")
            return False
        try:
            sol = self.hfss.odesign.GetModule("Solutions")
            header = ["IncludePortPostProcessing:=", False, "SpecifySystemPower:=", False]
            cmd = [header]
            for port, (pvar, phvar) in mapping.items():
                cmd.append(["Name:=", port, "Magnitude:=", pvar, "Phase:=", phvar])
            sol.EditSources(cmd)
            self.log_message(f"Solutions.EditSources applied to {len(mapping)} port(s).")
            return True
        except Exception as e:
            self.log_message(f"EditSources failed: {e}")
            return False

    def _create_infinite_sphere_after(self, name="Infinite Sphere1"):
        try:
            rf = self.hfss.odesign.GetModule("RadField")
            props = [
                f"NAME:{name}",
                "UseCustomRadiationSurface:=", False,
                "CSDefinition:=", "Theta-Phi",
                "Polarization:=", "Linear",
                "ThetaStart:=", "-108deg",
                "ThetaStop:=", "180deg",
                "ThetaStep:=", "1deg",
                "PhiStart:=", "-180deg",
                "PhiStop:=", "180deg",
                "PhiStep:=", "1deg",
                "UseLocalCS:=", False
            ]
            rf.InsertInfiniteSphereSetup(props)
            self.log_message(f"Infinite sphere '{name}' created.")
            return name
        except Exception as e:
            self.log_message(f"Infinite sphere creation failed: {e}")
            return None

    def _postprocess_after_solve(self):
        try:
            try:
                exs = self.hfss.get_excitations_name() or []
            except Exception:
                exs = []
            if not exs:
                exs = list(self.saved_excitations)
            if not exs:
                self.log_message("No excitations visible now; creating default vars p1/ph1..p4/ph4.")
                for i in range(1, 5):
                    self._ensure_post_var(f"p{i}", "1W")
                    self._ensure_post_var(f"ph{i}", "0deg")
                    self._ensure_post_var(f"Pow_P{i}", "1W")
                    self._ensure_post_var(f"Phi_P{i}", "0deg")
                return
            mapping = self._ensure_vars_for_ports(exs)
            self._apply_edit_sources(mapping)
            self._create_infinite_sphere_after("Infinite Sphere1")
        except Exception as e:
            self.log_message(f"Postprocess-after-solve error: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # ---------- Excitations UI actions ----------
    def refresh_ports_ui(self):
        exs = self._get_excitations_anyway()
        self.ports_info_label.configure(text=f"Ports: {len(exs)}")
        # remove linhas > 0
        for widget in self.ex_table.grid_slaves():
            info = widget.grid_info()
            if info.get("row", 0) >= 1:
                widget.grid_forget()
        self.ex_rows.clear()
        if not self.port_vars and exs:
            self._ensure_vars_for_ports(exs)
        for i, port in enumerate(exs, start=1):
            row = {}
            c0 = ctk.CTkLabel(self.ex_table, text=port)
            c0.grid(row=i, column=0, padx=10, pady=4, sticky="w")
            pvar, phvar = self.port_vars.get(port, (f"p{i}", f"ph{i}"))
            e1 = ctk.CTkEntry(self.ex_table, width=120); e1.insert(0, pvar)
            e1.grid(row=i, column=1, padx=10, pady=4); row["pvar"] = e1
            ev1 = ctk.CTkEntry(self.ex_table, width=120); ev1.insert(0, "1W")
            ev1.grid(row=i, column=2, padx=10, pady=4); row["pval"] = ev1
            e2 = ctk.CTkEntry(self.ex_table, width=120); e2.insert(0, phvar)
            e2.grid(row=i, column=3, padx=10, pady=4); row["phvar"] = e2
            ev2 = ctk.CTkEntry(self.ex_table, width=120); ev2.insert(0, "0deg")
            ev2.grid(row=i, column=4, padx=10, pady=4); row["phval"] = ev2
            self.ex_rows[port] = row

    def create_sync_post_vars(self):
        exs = self._get_excitations_anyway()
        if not exs:
            self.log_message("No ports to sync.")
            return
        mapping_ui: Dict[str, Tuple[str, str]] = {}
        for i, port in enumerate(exs, start=1):
            row = self.ex_rows.get(port)
            if not row:
                continue
            pvar = row["pvar"].get().strip() or f"p{i}"
            phvar = row["phvar"].get().strip() or f"ph{i}"
            pval = row["pval"].get().strip() or "1W"
            phval = row["phval"].get().strip() or "0deg"
            self._ensure_post_var(pvar, pval)
            self._ensure_post_var(phvar, phval)
            idx = self._port_index_from_name(port) or i
            self._ensure_post_var(f"Pow_P{idx}", pval)
            self._ensure_post_var(f"Phi_P{idx}", phval)
            mapping_ui[port] = (pvar, phvar)
        self.port_vars = mapping_ui
        self.log_message("Post variables created/synced for all ports.")

    def apply_sources_from_vars(self):
        if not self.port_vars:
            exs = self._get_excitations_anyway()
            if not exs:
                self.log_message("No ports available to apply sources.")
                return
            self.port_vars = self._ensure_vars_for_ports(exs)
        self._apply_edit_sources(self.port_vars)

    def set_equal_power_1w(self):
        for row in self.ex_rows.values():
            row["pval"].delete(0, "end"); row["pval"].insert(0, "1W")
        self.log_message("All power values set to 1W (UI only).")

    def zero_all_phases(self):
        for row in self.ex_rows.values():
            row["phval"].delete(0, "end"); row["phval"].insert(0, "0deg")
        self.log_message("All phases set to 0deg (UI only).")

    def update_var_values_in_aedt(self):
        for row in self.ex_rows.values():
            pvar = row["pvar"].get().strip()
            phvar = row["phvar"].get().strip()
            pval = row["pval"].get().strip()
            phval = row["phval"].get().strip()
            if pvar and pval: self._ensure_post_var(pvar, pval)
            if phvar and phval: self._ensure_post_var(phvar, phval)
        self.log_message("Variable values updated in AEDT (no EditSources).")

    # util
    def _get_excitations_anyway(self) -> List[str]:
        exs = []
        try:
            exs = self.hfss.get_excitations_name() or []
        except Exception:
            pass
        if not exs:
            exs = list(self.saved_excitations)
        if not exs:
            try:
                exs = list(getattr(self.hfss, "excitations", []) or [])
            except Exception:
                exs = []
        return exs

    # --------- Result analysis (S11 + tuning) ----------
    def _get_context_candidates(self) -> List[str]:
        return ["Setup1 : Sweep1", "Setup1: Sweep1", "Setup1 : Sweep", "Setup1: Sweep"]

    def _fetch_solution(self, expression: str):
        """Tenta obter solução de forma robusta para um expression."""
        last_err = None
        for ctx in self._get_context_candidates():
            try:
                rpt = self.hfss.post.reports_by_category.standard(expressions=[expression], setup=ctx)
                sol = rpt.get_solution_data()
                if sol and hasattr(sol, "primary_sweep_values"):
                    return sol, ctx
            except Exception as e:
                last_err = e
        if last_err:
            self.log_message(f"Report fetch failed for '{expression}': {last_err}")
        return None, None

    def _get_s11_curves(self, port_expr: str):
        """Retorna (f_ghz, s11_db, reS, imS) para o port_expr='(P1_Lumped,P1_Lumped)'."""
        # dB curve
        sol_db, ctx = self._fetch_solution(f"dB(S{port_expr})")
        if not sol_db:
            return None
        try:
            f_db = np.asarray(sol_db.primary_sweep_values, dtype=float)
            y_db = np.asarray(sol_db.data_real()[0], dtype=float)
        except Exception:
            return None
        # Re/Im for S11
        sol_re, _ = self._fetch_solution(f"re(S{port_expr})")
        sol_im, _ = self._fetch_solution(f"im(S{port_expr})")
        if not sol_re or not sol_im:
            # fallback: sem parte complexa
            return (f_db, y_db, None, None)
        try:
            f_re = np.asarray(sol_re.primary_sweep_values, dtype=float)
            f_im = np.asarray(sol_im.primary_sweep_values, dtype=float)
            reS = np.asarray(sol_re.data_real()[0], dtype=float)
            imS = np.asarray(sol_im.data_real()[0], dtype=float)
            # alinhar ao mesmo grid do dB (assumindo igual)
            return (f_db, y_db, reS, imS)
        except Exception:
            return (f_db, y_db, None, None)

    @staticmethod
    def _nearest_index(vec: np.ndarray, value: float) -> int:
        return int(np.argmin(np.abs(vec - value)))

    def analyze_and_mark_s11(self):
        """Acha o menor pico de S11, calcula Z no ponto e marca no gráfico."""
        try:
            exs = self._get_excitations_anyway()
            if not exs:
                self.log_message("Analyze S11: no excitations found.")
                return
            # usar a porta de menor índice
            ex_sorted = sorted(exs, key=lambda p: self._port_index_from_name(p) or 9999)
            port_name = ex_sorted[0]
            pshort = port_name.split(":")[0]  # 'P1_Lumped'
            port_expr = f"({pshort},{pshort})"

            data = self._get_s11_curves(port_expr)
            if not data:
                self.log_message("No S11 data available for analysis.")
                return
            f, s11_db, reS, imS = data
            if f.size == 0:
                self.log_message("S11 curve is empty.")
                return

            idx_min = int(np.argmin(s11_db))
            f_res = float(f[idx_min])
            s11_min_db = float(s11_db[idx_min])

            # calcular Z no ponto usando S complexo
            R = X = None
            if reS is not None and imS is not None and len(reS) == len(f):
                S = complex(reS[idx_min], imS[idx_min])
                Z0 = 50.0
                try:
                    Z = Z0 * (1 + S) / (1 - S)
                    R = float(np.real(Z))
                    X = float(np.imag(Z))
                except ZeroDivisionError:
                    R, X = None, None

            # atualizar plot
            self.ax_s11.clear()
            self.ax_s11.grid(True, alpha=0.5)
            self.ax_s11.set_xlabel("Frequency (GHz)")
            self.ax_s11.set_ylabel("S11 (dB)")
            self.ax_s11.set_title(f"S11 - {pshort}")
            self.ax_s11.plot(f, s11_db, linewidth=2, label="S11 (dB)")
            self.ax_s11.axhline(-10, linestyle="--", alpha=0.6, label="-10 dB")
            self.ax_s11.axvline(f_res, linestyle="--", alpha=0.6, label=f"Res @ {f_res:.4g} GHz")
            y_min = s11_min_db
            self.ax_s11.scatter([f_res], [y_min], s=40)

            # anotação de impedância
            if R is not None and X is not None:
                txt = f"Z ≈ {R:.1f} {'+' if X>=0 else '–'} j{abs(X):.1f} Ω"
            else:
                txt = f"S11 min = {s11_min_db:.2f} dB"
            self.ax_s11.annotate(txt, xy=(f_res, y_min),
                                 xytext=(f_res, y_min + 5),
                                 arrowprops=dict(arrowstyle="->", alpha=0.6))

            self.ax_s11.legend()
            self.fig.tight_layout()
            self.canvas.draw()

            # escala sugerida (regra direta em f): s = f_res / f_target
            f_target = float(self.params["frequency"])
            s = f_res / f_target if f_target > 0 else 1.0
            self.last_analysis.update({
                "port": pshort,
                "f_res": f_res,
                "s11_min_db": s11_min_db,
                "R": R, "X": X,
                "scale": s
            })
            dir_txt = "increase (bigger)" if s > 1 else "decrease (smaller)"
            self.analysis_label.configure(
                text=(f"Resonance: {f_res:.4g} GHz | Target: {f_target:.4g} GHz | "
                      f"Suggested scale s = f_res/f_target = {s:.5f} → {dir_txt}\n"
                      f"Z@res: {txt}")
            )
            self.log_message(f"S11 analyzed. f_res={f_res:.6g} GHz, S11min={s11_min_db:.2f} dB, scale={s:.6f}")
        except Exception as e:
            self.log_message(f"Analyze S11 error: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # helpers para variáveis (mm)
    def _get_var_mm(self, name: str) -> Optional[float]:
        try:
            val = self.hfss[name]  # string tipo '12.3mm'
            if isinstance(val, (int, float)):
                return float(val)
            s = str(val).strip().lower().replace(" ", "")
            if s.endswith("mm"):
                return float(s[:-2])
            # se vier em outras unidades, deixa como está (assumir mm se numérico)
            return float(s)
        except Exception:
            return None

    def _set_var_mm(self, name: str, value_mm: float):
        try:
            self.hfss[name] = f"{value_mm}mm"
            return True
        except Exception as e:
            self.log_message(f"Set var '{name}' failed: {e}")
            return False

    def apply_resonance_correction(self):
        """Aplica correção geométrica usando scale s = f_res / f_target."""
        try:
            f_res = self.last_analysis.get("f_res")
            s = self.last_analysis.get("scale")
            if f_res is None or s is None:
                self.log_message("Apply correction: run 'Analyze S11' first.")
                return
            f_target = float(self.params["frequency"])

            # limitações suaves (evitar escalas absurdas)
            s_clamped = max(0.6, min(1.4, s))
            if abs(s_clamped - s) > 1e-6:
                self.log_message(f"Scale clamped from {s:.4f} to {s_clamped:.4f}")
            s = s_clamped

            scale_vars = ["patchL", "patchW", "spacing", "subW", "subL"]
            ok = True
            for v in scale_vars:
                cur = self._get_var_mm(v)
                if cur is None:
                    self.log_message(f"Cannot read var '{v}'. Skipping.")
                    ok = False
                    continue
                newv = cur * s
                if not self._set_var_mm(v, newv):
                    ok = False
                else:
                    self.log_message(f"{v}: {cur:.4f} mm → {newv:.4f} mm (s={s:.6f})")

            if ok:
                self.analysis_label.configure(
                    text=(self.analysis_label.cget("text") +
                          f"\nApplied scale s={s:.6f} to variables {scale_vars}. Re-run the simulation.")
                )
                self.log_message("Resonance correction applied. Please re-run the simulation.")
            else:
                self.log_message("Resonance correction partially applied (see messages).")
        except Exception as e:
            self.log_message(f"Apply correction error: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # ------------- Simulação -------------
    def start_simulation_thread(self):
        if self.is_simulation_running:
            self.log_message("Simulation is already running")
            return
        self.stop_simulation = False
        self.is_simulation_running = True
        threading.Thread(target=self.run_simulation, daemon=True).start()

    def stop_simulation_thread(self):
        self.stop_simulation = True
        self.log_message("Simulation stop requested")

    def run_simulation(self):
        try:
            self.log_message("Starting simulation")
            self.run_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            self.sim_status_label.configure(text="Simulation in progress")
            self.progress_bar.set(0)

            if not self.get_parameters():
                self.log_message("Invalid parameters. Aborting.")
                return
            if self.calculated_params["num_patches"] < 1:
                self.calculate_parameters()

            self._open_or_create_project()
            self.progress_bar.set(0.20)

            self.hfss.modeler_model_units = "mm"
            self.log_message("Model units set to: mm")

            sub_name = self.params["substrate_material"]
            if not self.hfss.materials.checkifmaterialexists(sub_name):
                sub_name = "Custom_Substrate"
                self._ensure_material(sub_name, float(self.params["er"]), float(self.params["tan_d"]))

            L = float(self.calculated_params["patch_length"])
            W = float(self.calculated_params["patch_width"])
            spacing = float(self.calculated_params["spacing"])
            rows = int(self.calculated_params["rows"])
            cols = int(self.calculated_params["cols"])
            h_sub = float(self.params["substrate_thickness"])
            sub_w = float(self.calculated_params["substrate_width"])
            sub_l = float(self.calculated_params["substrate_length"])

            self._set_design_variables(L, W, spacing, rows, cols, h_sub, sub_w, sub_l)

            # geometria
            self.log_message("Creating substrate")
            substrate = self.hfss.modeler.create_box(
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL", "h_sub"],
                name="Substrate",
                material=sub_name
            )
            self.log_message("Creating ground plane")
            ground = self.hfss.modeler.create_rectangle(
                orientation="XY",
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL"],
                name="Ground",
                material="copper"
            )

            self.log_message(f"Creating {rows*cols} patches in {rows}x{cols} configuration")
            patches: List = []
            total_w = cols * W + (cols - 1) * spacing
            total_l = rows * L + (rows - 1) * spacing
            start_x = -total_w / 2 + W / 2
            start_y = -total_l / 2 + L / 2
            self.progress_bar.set(0.30)

            count = 0
            for r in range(rows):
                for c in range(cols):
                    if self.stop_simulation:
                        self.log_message("Simulation stopped by user")
                        return
                    count += 1
                    patch_name = f"Patch_{count}"
                    cx = start_x + c * (W + spacing)
                    cy = start_y + r * (L + spacing)

                    origin = [cx - W / 2, cy - L / 2, "h_sub"]
                    self.log_message(f"Creating patch {count} at ({r}, {c})")

                    patch = self.hfss.modeler.create_rectangle(
                        orientation="XY",
                        origin=origin,
                        sizes=["patchW", "patchL"],
                        name=patch_name,
                        material="copper"
                    )
                    patches.append(patch)

                    if self.params["feed_position"] == "edge":
                        y_feed = cy - 0.5 * L + 0.02 * L
                    else:
                        y_feed = cy - 0.5 * L + 0.30 * L
                    relx = min(max(float(self.params["feed_rel_x"]), 0.0), 1.0)
                    x_feed = cx - 0.5 * W + relx * W

                    pad = self.hfss.modeler.create_circle(
                        orientation="XY",
                        origin=[x_feed, y_feed, "h_sub"],
                        radius="a",
                        name=f"{patch_name}_Pad",
                        material="copper"
                    )
                    try:
                        self.hfss.modeler.unite([patch, pad])
                    except Exception:
                        pass

                    self._create_coax_feed_lumped(
                        ground=ground,
                        substrate=substrate,
                        x_feed=x_feed,
                        y_feed=y_feed,
                        name_prefix=f"P{count}"
                    )
                    self.progress_bar.set(0.30 + 0.25 * (count / float(rows * cols)))

            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return

            try:
                names = [ground.name] + [p.name for p in patches]
                self.hfss.assign_perfecte_to_sheets(names)
                self.log_message(f"PerfectE assigned to: {names}")
            except Exception as e:
                self.log_message(f"PerfectE assignment warning: {e}")

            self.log_message("Creating air region + radiation boundary")
            lambda0_mm = self.c / (self.params["sweep_start"] * 1e9) * 1000.0
            pad_mm = float(lambda0_mm) / 4.0
            region = self.hfss.modeler.create_region(
                [pad_mm, pad_mm, pad_mm, pad_mm, pad_mm, pad_mm], is_percentage=False
            )
            self.hfss.assign_radiation_boundary_to_objects(region)
            self.progress_bar.set(0.60)

            self.log_message("Creating simulation setup")
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02

            self.log_message(f"Creating frequency sweep: {self.params['sweep_type']}")
            stype = self.params["sweep_type"]
            try:
                try:
                    sw = setup.get_sweep("Sweep1")
                    if sw:
                        sw.delete()
                except Exception:
                    pass
                if stype == "Discrete":
                    step = float(self.params["sweep_step"])
                    setup.create_linear_step_sweep(
                        unit="GHz",
                        start_frequency=self.params["sweep_start"],
                        stop_frequency=self.params["sweep_stop"],
                        step_size=step,
                        name="Sweep1"
                    )
                elif stype == "Fast":
                    setup.create_frequency_sweep(
                        unit="GHz",
                        name="Sweep1",
                        start_frequency=self.params["sweep_start"],
                        stop_frequency=self.params["sweep_stop"],
                        sweep_type="Fast"
                    )
                else:
                    setup.create_frequency_sweep(
                        unit="GHz",
                        name="Sweep1",
                        start_frequency=self.params["sweep_start"],
                        stop_frequency=self.params["sweep_stop"],
                        sweep_type="Interpolating"
                    )
            except Exception as e:
                self.log_message(f"Sweep creation warning: {e}")

            # malha leve
            try:
                lambda_g_mm = max(1e-6, self.calculated_params["lambda_g"])
                edge_len = max(lambda_g_mm / 60.0, W / 200.0)
                for p in patches:
                    self.hfss.mesh.assign_length_mesh([p], maximum_length=f"{edge_len}mm")
            except Exception as e:
                self.log_message(f"Mesh refinement warning: {e}")

            # excitations -> salvar
            exs = self._get_excitations_anyway()
            self.saved_excitations = list(exs)
            self.log_message(f"Excitations created: {len(exs)} -> {exs}")

            self.log_message("Validating design")
            try:
                _ = self.hfss.validate_full_design()
            except Exception as e:
                self.log_message(f"Validation warning: {e}")

            self.log_message("Starting analysis")
            if self.save_project:
                self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])

            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return

            # Pós-processamento (variáveis + esfera)
            self._postprocess_after_solve()

            self.progress_bar.set(0.95)
            self.sim_status_label.configure(text="Simulation completed")
            self.log_message("Simulation completed successfully")

            # Atualiza UI Excitations
            self.refresh_ports_ui()

            # Plot/análise automática
            if self.params.get("plot_results", True):
                self.analyze_and_mark_s11()
            else:
                self.log_message("Plots skipped (plot_results=False).")

            self.progress_bar.set(1.0)
        except Exception as e:
            self.log_message(f"Error in simulation: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            self.sim_status_label.configure(text=f"Simulation error: {e}")
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    # ------------- Encerramento -------------
    def cleanup(self):
        try:
            if self.hfss:
                try:
                    if self.save_project:
                        self.hfss.save_project()
                    else:
                        self.hfss.close_project(save=False)
                except Exception as e:
                    self.log_message(f"Error closing project: {e}")
            if self.desktop:
                try:
                    self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception as e:
                    self.log_message(f"Error releasing desktop: {e}")
            if self.temp_folder and not self.save_project:
                try:
                    self.temp_folder.cleanup()
                except Exception as e:
                    self.log_message(f"Error cleaning up temporary files: {e}")
        except Exception as e:
            self.log_message(f"Error during cleanup: {e}")

    def on_closing(self):
        self.log_message("Application closing...")
        self.cleanup()
        self.window.quit()
        self.window.destroy()

    def save_parameters(self):
        try:
            all_params = {**self.params, **self.calculated_params}
            with open("antenna_parameters.json", "w") as f:
                json.dump(all_params, f, indent=4)
            self.log_message("Parameters saved to antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error saving parameters: {e}")

    def load_parameters(self):
        try:
            with open("antenna_parameters.json", "r") as f:
                all_params = json.load(f)
            for k in self.params:
                if k in all_params:
                    self.params[k] = all_params[k]
            for k in self.calculated_params:
                if k in all_params:
                    self.calculated_params[k] = all_params[k]
            self.update_interface_from_params()
            self.log_message("Parameters loaded from antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error loading parameters: {e}")

    def update_interface_from_params(self):
        try:
            for key, widget in self.entries:
                if key in self.params:
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end")
                        widget.insert(0, str(self.params[key]))
                    elif isinstance(widget, ctk.StringVar):
                        widget.set(self.params[key])
                    elif isinstance(widget, ctk.BooleanVar):
                        widget.set(self.params[key])
            self.patches_label.configure(text=f"Number of Patches: {self.calculated_params['num_patches']}")
            self.rows_cols_label.configure(text=f"Configuration: {self.calculated_params['rows']} x "
                                                f"{self.calculated_params['cols']}")
            self.spacing_label.configure(text=f"Spacing: {self.calculated_params['spacing']:.2f} mm "
                                              f"({self.params['spacing_type']})")
            self.dimensions_label.configure(
                text=f"Patch Dimensions: {self.calculated_params['patch_length']:.2f} x "
                     f"{self.calculated_params['patch_width']:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {self.calculated_params['lambda_g']:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset (y): {self.calculated_params['feed_offset']:.2f} mm")
            self.substrate_dims_label.configure(
                text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x "
                     f"{self.calculated_params['substrate_length']:.2f} mm")
            self.log_message("Interface updated with loaded parameters")
        except Exception as e:
            self.log_message(f"Error updating interface: {e}")

    def run(self):
        try:
            self.window.mainloop()
        finally:
            self.cleanup()


if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()
