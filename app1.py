# -*- coding: utf-8 -*-
import os
import tempfile
from datetime import datetime
import math
import json
import traceback
import queue
import threading
from typing import Tuple, List, Optional, Dict

import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401 (needed by mplot3d)
import customtkinter as ctk

from ansys.aedt.core import Desktop, Hfss

# ---------- Appearance ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ModernPatchAntennaDesigner:
    def __init__(self):
        # AEDT
        self.hfss: Optional[Hfss] = None
        self.desktop: Optional[Desktop] = None
        self.temp_folder = None
        self.project_path = ""
        self.project_display_name = "patch_array"
        self.design_base_name = "patch_array"

        # Runtime
        self.log_queue = queue.Queue()
        self.is_simulation_running = False
        self.save_project = False
        self.stop_simulation = False
        self.created_ports: List[str] = []

        # Data
        self.last_s11_analysis = None  # dict f, s11_db, |Z|, f_res, Z(f_res)...
        self.theta_cut = None          # (theta, gain)
        self.phi_cut = None            # (phi, gain)
        self.grid3d = None             # (TH, PH, Gdb)
        self.auto_refresh_job = None

        # User params
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
            "sweep_type": "Interpolating",  # "Discrete" | "Interpolating" | "Fast"
            "sweep_step": 0.02,            # GHz (Discrete)
            # 3D sampling
            "theta_step": 10.0,            # deg
            "phi_step": 10.0               # deg
        }

        # Calculated params
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
        self.setup_gui()

    # ---------------- GUI ----------------
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        try:
            self.window.state("zoomed")
        except Exception:
            self.window.geometry("1500x950")

        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.window, height=70, fg_color=("gray85", "gray20"))
        header.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header.grid_propagate(False)
        ctk.CTkLabel(
            header, text="Patch Antenna Array Designer",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=("gray10", "gray90")
        ).pack(pady=18)

        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        for name in ["Design Parameters", "Simulation", "Results", "Log"]:
            self.tabview.add(name)
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()

        status = ctk.CTkFrame(self.window, height=38)
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        self.status_label = ctk.CTkLabel(status, text="Ready", font=ctk.CTkFont(weight="bold"))
        self.status_label.pack(pady=6)

        self.process_log_queue()

    def create_section(self, parent, title, row, column, padx=10, pady=10):
        section = ctk.CTkFrame(parent, fg_color=("gray92", "gray18"))
        section.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
        section.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(section, text=title, font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=("gray20", "gray80")).grid(row=0, column=0, sticky="w", padx=15, pady=(10, 6))
        ctk.CTkFrame(section, height=2, fg_color=("gray70", "gray30")).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        return section

    def setup_parameters_tab(self):
        tab = self.tabview.tab("Design Parameters")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkScrollableFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main.grid_columnconfigure(0, weight=1)

        sec_ant = self.create_section(main, "Antenna Parameters", 0, 0)
        self.entries = []
        row_idx = 2

        def add_entry(section, label, key, value, row, combo=None, check=False):
            ctk.CTkLabel(section, text=label, font=ctk.CTkFont(weight="bold")
                         ).grid(row=row, column=0, padx=15, pady=6, sticky="w")
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(section, values=combo, variable=var, width=220)
                widget.grid(row=row, column=1, padx=15, pady=6)
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(section, text="", variable=var)
                widget.grid(row=row, column=1, padx=15, pady=6, sticky="w")
                self.entries.append((key, var))
            else:
                widget = ctk.CTkEntry(section, width=220)
                widget.insert(0, str(value))
                widget.grid(row=row, column=1, padx=15, pady=6)
                self.entries.append((key, widget))
            return row + 1

        row_idx = add_entry(sec_ant, "Central Frequency (GHz):", "frequency", self.params["frequency"], row_idx)
        row_idx = add_entry(sec_ant, "Desired Gain (dBi):", "gain", self.params["gain"], row_idx)
        row_idx = add_entry(sec_ant, "Sweep Start (GHz):", "sweep_start", self.params["sweep_start"], row_idx)
        row_idx = add_entry(sec_ant, "Sweep Stop (GHz):", "sweep_stop", self.params["sweep_stop"], row_idx)
        row_idx = add_entry(sec_ant, "Patch Spacing:", "spacing_type", self.params["spacing_type"], row_idx,
                            combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        sec_sub = self.create_section(main, "Substrate Parameters", 1, 0)
        row_idx = 2
        row_idx = add_entry(sec_sub, "Substrate Material:", "substrate_material",
                            self.params["substrate_material"], row_idx,
                            combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        row_idx = add_entry(sec_sub, "Relative Permittivity (εr):", "er", self.params["er"], row_idx)
        row_idx = add_entry(sec_sub, "Loss Tangent (tan δ):", "tan_d", self.params["tan_d"], row_idx)
        row_idx = add_entry(sec_sub, "Substrate Thickness (mm):", "substrate_thickness", self.params["substrate_thickness"], row_idx)
        row_idx = add_entry(sec_sub, "Metal Thickness (mm):", "metal_thickness", self.params["metal_thickness"], row_idx)

        sec_coax = self.create_section(main, "Coaxial Feed Parameters", 2, 0)
        row_idx = 2
        row_idx = add_entry(sec_coax, "Feed position type:", "feed_position", self.params["feed_position"], row_idx,
                            combo=["inset", "edge"])
        row_idx = add_entry(sec_coax, "Feed relative X (0..1):", "feed_rel_x", self.params["feed_rel_x"], row_idx)
        row_idx = add_entry(sec_coax, "Inner radius a (mm):", "probe_radius", self.params["probe_radius"], row_idx)
        row_idx = add_entry(sec_coax, "b/a ratio:", "coax_ba_ratio", self.params["coax_ba_ratio"], row_idx)
        row_idx = add_entry(sec_coax, "Shield wall (mm):", "coax_wall_thickness", self.params["coax_wall_thickness"], row_idx)
        row_idx = add_entry(sec_coax, "Port length below GND Lp (mm):", "coax_port_length", self.params["coax_port_length"], row_idx)
        row_idx = add_entry(sec_coax, "Anti-pad clearance (mm):", "antipad_clearance", self.params["antipad_clearance"], row_idx)

        sec_sim = self.create_section(main, "Simulation Settings", 3, 0)
        row_idx = 2
        row_idx = add_entry(sec_sim, "CPU Cores:", "cores", self.params["cores"], row_idx)
        row_idx = add_entry(sec_sim, "Show HFSS Interface:", "show_gui", not self.params["non_graphical"], row_idx, check=True)
        row_idx = add_entry(sec_sim, "Save Project:", "save_project", self.save_project, row_idx, check=True)
        row_idx = add_entry(sec_sim, "Sweep Type:", "sweep_type", self.params["sweep_type"], row_idx,
                            combo=["Discrete", "Interpolating", "Fast"])
        row_idx = add_entry(sec_sim, "Discrete Step (GHz):", "sweep_step", self.params["sweep_step"], row_idx)
        row_idx = add_entry(sec_sim, "3D Theta step (deg):", "theta_step", self.params["theta_step"], row_idx)
        row_idx = add_entry(sec_sim, "3D Phi step (deg):", "phi_step", self.params["phi_step"], row_idx)

        sec_calc = self.create_section(main, "Calculated Parameters", 4, 0)
        grid = ctk.CTkFrame(sec_calc); grid.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        grid.columnconfigure((0, 1), weight=1)
        self.patches_label = ctk.CTkLabel(grid, text="Number of Patches: 4", font=ctk.CTkFont(weight="bold"))
        self.patches_label.grid(row=0, column=0, sticky="w", pady=4)
        self.rows_cols_label = ctk.CTkLabel(grid, text="Configuration: 2 x 2", font=ctk.CTkFont(weight="bold"))
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=4)
        self.spacing_label = ctk.CTkLabel(grid, text="Spacing: -- mm", font=ctk.CTkFont(weight="bold"))
        self.spacing_label.grid(row=1, column=0, sticky="w", pady=4)
        self.dimensions_label = ctk.CTkLabel(grid, text="Patch Dimensions: -- x -- mm", font=ctk.CTkFont(weight="bold"))
        self.dimensions_label.grid(row=1, column=1, sticky="w", pady=4)
        self.lambda_label = ctk.CTkLabel(grid, text="Guided Wavelength: -- mm", font=ctk.CTkFont(weight="bold"))
        self.lambda_label.grid(row=2, column=0, sticky="w", pady=4)
        self.feed_offset_label = ctk.CTkLabel(grid, text="Feed Offset (y): -- mm", font=ctk.CTkFont(weight="bold"))
        self.feed_offset_label.grid(row=2, column=1, sticky="w", pady=4)
        self.substrate_dims_label = ctk.CTkLabel(grid, text="Substrate Dimensions: -- x -- mm",
                                                 font=ctk.CTkFont(weight="bold"))
        self.substrate_dims_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=4)

        btns = ctk.CTkFrame(sec_calc); btns.grid(row=3, column=0, sticky="ew", padx=15, pady=12)
        ctk.CTkButton(btns, text="Calculate Parameters", command=self.calculate_parameters,
                      fg_color="#2E8B57", hover_color="#3CB371", width=180).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="Save Parameters", command=self.save_parameters,
                      fg_color="#4169E1", hover_color="#6495ED", width=140).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="Load Parameters", command=self.load_parameters,
                      fg_color="#FF8C00", hover_color="#FFA500", width=140).pack(side="left", padx=8)

    def setup_simulation_tab(self):
        tab = self.tabview.tab("Simulation")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(main, text="Simulation Control", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        row = ctk.CTkFrame(main); row.pack(pady=14)
        self.run_button = ctk.CTkButton(row, text="Run Simulation", command=self.start_simulation_thread,
                                        fg_color="#2E8B57", hover_color="#3CB371", height=40, width=160)
        self.run_button.pack(side="left", padx=8)
        self.stop_button = ctk.CTkButton(row, text="Stop Simulation", command=self.stop_simulation_thread,
                                         fg_color="#DC143C", hover_color="#FF4500",
                                         state="disabled", height=40, width=160)
        self.stop_button.pack(side="left", padx=8)
        bar = ctk.CTkFrame(main); bar.pack(fill="x", padx=50, pady=8)
        ctk.CTkLabel(bar, text="Simulation Progress:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.progress_bar = ctk.CTkProgressBar(bar, height=18); self.progress_bar.pack(fill="x", pady=6); self.progress_bar.set(0)
        self.sim_status_label = ctk.CTkLabel(main, text="Simulation not started", font=ctk.CTkFont(weight="bold"))
        self.sim_status_label.pack(pady=8)
        note = ctk.CTkFrame(main, fg_color=("gray90", "gray15")); note.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(note, text="Tip: With post vars (p_i / ph_i) you can retune beams without re-solving.",
                     font=ctk.CTkFont(size=12, slant="italic"), text_color=("gray40", "gray60")).pack(padx=10, pady=10)

    def setup_results_tab(self):
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(main, text="Results & Beamforming", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=10)

        # Graph area with GridSpec 3x2
        graph_frame = ctk.CTkFrame(main)
        graph_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)

        self.fig = plt.figure(figsize=(12, 10))
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        self.fig.patch.set_facecolor(face)
        gs = self.fig.add_gridspec(3, 2, height_ratios=[1, 1, 1.3])

        self.ax_s11 = self.fig.add_subplot(gs[0, 0])
        self.ax_imp = self.fig.add_subplot(gs[0, 1])
        self.ax_th = self.fig.add_subplot(gs[1, 0])
        self.ax_ph = self.fig.add_subplot(gs[1, 1])
        self.ax_3d = self.fig.add_subplot(gs[2, :], projection='3d')

        for ax in (self.ax_s11, self.ax_imp, self.ax_th, self.ax_ph):
            ax.set_facecolor(face)
            if ctk.get_appearance_mode() == "Dark":
                ax.tick_params(colors='white')
                ax.xaxis.label.set_color('white'); ax.yaxis.label.set_color('white'); ax.title.set_color('white')
                for s in ['bottom', 'top', 'right', 'left']: ax.spines[s].set_color('white')
                ax.grid(color='gray', alpha=0.5)
        self.ax_3d.set_facecolor(face)

        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # Beamforming / refresh panel
        panel = ctk.CTkFrame(main)
        panel.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        panel.grid_columnconfigure(0, weight=1)
        self.src_frame = ctk.CTkFrame(panel); self.src_frame.grid(row=0, column=0, sticky="ew", padx=6, pady=(8, 2))
        self.source_controls: Dict[str, Dict[str, ctk.CTkBaseClass]] = {}
        self.auto_refresh_var = ctk.BooleanVar(value=False)

        ctrl = ctk.CTkFrame(panel); ctrl.grid(row=1, column=0, pady=6)
        ctk.CTkButton(ctrl, text="Analyze S11", command=self.analyze_and_mark_s11,
                      fg_color="#6A5ACD", hover_color="#7B68EE").pack(side="left", padx=8)
        ctk.CTkButton(ctrl, text="Apply Sources", command=self.apply_sources_from_ui,
                      fg_color="#20B2AA", hover_color="#40E0D0").pack(side="left", padx=8)
        ctk.CTkButton(ctrl, text="Refresh Patterns", command=self.refresh_patterns_only,
                      fg_color="#FF8C00", hover_color="#FFA500").pack(side="left", padx=8)
        ctk.CTkCheckBox(ctrl, text="Auto-refresh (1.5s)", variable=self.auto_refresh_var,
                        command=self.toggle_auto_refresh).pack(side="left", padx=8)
        ctk.CTkButton(ctrl, text="Export PNG", command=self.export_png,
                      fg_color="#20B2AA", hover_color="#40E0D0").pack(side="left", padx=8)
        ctk.CTkButton(ctrl, text="Export CSV (S11)", command=self.export_csv,
                      fg_color="#6A5ACD", hover_color="#7B68EE").pack(side="left", padx=8)

        self.result_label = ctk.CTkLabel(main, text="", font=ctk.CTkFont(weight="bold"))
        self.result_label.grid(row=3, column=0, pady=(6, 2))

    def setup_log_tab(self):
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(main, text="Simulation Log", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=10)
        self.log_text = ctk.CTkTextbox(main, width=900, height=500, font=ctk.CTkFont(family="Consolas"))
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        btn = ctk.CTkFrame(main); btn.grid(row=2, column=0, pady=8)
        ctk.CTkButton(btn, text="Clear Log", command=self.clear_log).pack(side="left", padx=8)
        ctk.CTkButton(btn, text="Save Log", command=self.save_log).pack(side="left", padx=8)

    # ------------- Log utilities -------------
    def log_message(self, message):
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")

    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert("end", msg)
                self.log_text.see("end")
        except queue.Empty:
            pass
        finally:
            if self.window.winfo_exists():
                self.window.after(100, self.process_log_queue)

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

    # ----------- Physics / calculations -----------
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
                             "sweep_step", "theta_step", "phi_step"]:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
                elif key in ["spacing_type", "substrate_material", "feed_position", "sweep_type"]:
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
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)

    def _size_array_from_gain(self):
        G_elem = 8.0
        G_des = float(self.params["gain"])
        N_req = max(1, int(math.ceil(10 ** ((G_des - G_elem) / 10.0))))
        if N_req % 2 == 1: N_req += 1
        rows = max(2, int(round(math.sqrt(N_req))));  rows += rows % 2
        cols = max(2, int(math.ceil(N_req / rows))); cols += cols % 2
        while rows * cols < N_req:
            if rows <= cols: rows += 2
            else: cols += 2
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
            self.log_message(f"Array sizing -> target gain {self.params['gain']} dBi, N_req≈{N_req}, layout {rows}x{cols} (= {rows*cols} patches)")
            self.calculated_params["feed_offset"] = 0.30 * L_mm
            self.calculate_substrate_size()
            # UI
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
            self.log_message(f"Error in calculation: {e}\nTraceback: {traceback.format_exc()}")

    # --------- AEDT helpers ---------
    def _ensure_material(self, name: str, er: float, tan_d: float):
        try:
            if not self.hfss.materials.checkifmaterialexists(name):
                self.hfss.materials.add_material(name)
                m = self.hfss.materials.material_keys[name]
                m.permittivity = er; m.dielectric_loss_tangent = tan_d
                self.log_message(f"Created material: {name} (er={er}, tanδ={tan_d})")
        except Exception as e:
            self.log_message(f"Material management warning for '{name}': {e}")

    def _open_or_create_project(self):
        if self.desktop is None:
            self.desktop = Desktop(version=self.params["aedt_version"],
                                   non_graphical=self.params["non_graphical"], new_desktop=True)
        if self.temp_folder is None:
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.project_path = os.path.join(self.temp_folder.name, f"{self.project_display_name}_{ts}.aedt")
        self.hfss = Hfss(project=self.project_path, design=self.design_base_name, solution_type="DrivenModal",
                         version=self.params["aedt_version"], non_graphical=self.params["non_graphical"])
        self.log_message(f"Created new project: {self.project_path} (design '{self.design_base_name}')")

    def _set_design_variables(self, L, W, spacing, rows, cols, h_sub, sub_w, sub_l):
        a = float(self.params["probe_radius"]); ba = float(self.params["coax_ba_ratio"])
        b = a * ba; wall = float(self.params["coax_wall_thickness"]); Lp = float(self.params["coax_port_length"])
        clear = float(self.params["antipad_clearance"])
        self.hfss["f0"] = f"{self.params['frequency']}GHz"
        self.hfss["h_sub"] = f"{h_sub}mm"; self.hfss["t_met"] = f"{self.params['metal_thickness']}mm"
        self.hfss["patchL"] = f"{L}mm"; self.hfss["patchW"] = f"{W}mm"
        self.hfss["spacing"] = f"{spacing}mm"; self.hfss["rows"] = str(rows); self.hfss["cols"] = str(cols)
        self.hfss["subW"] = f"{sub_w}mm"; self.hfss["subL"] = f"{sub_l}mm"
        self.hfss["a"] = f"{a}mm"; self.hfss["b"] = f"{b}mm"; self.hfss["wall"] = f"{wall}mm"
        self.hfss["Lp"] = f"{Lp}mm"; self.hfss["clear"] = f"{clear}mm"; self.hfss["eps"] = "0.001mm"
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
            if b_val - a_val < 0.02: b_val = a_val + 0.02

            pin = self.hfss.modeler.create_cylinder("Z", [x_feed, y_feed, -Lp_val],
                                                    a_val, h_sub_val + Lp_val + 0.001,
                                                    name=f"{name_prefix}_Pin", material="copper")
            shield_outer = self.hfss.modeler.create_cylinder("Z", [x_feed, y_feed, -Lp_val],
                                                             b_val + wall_val, Lp_val,
                                                             name=f"{name_prefix}_ShieldOuter", material="copper")
            shield_inner_void = self.hfss.modeler.create_cylinder("Z", [x_feed, y_feed, -Lp_val],
                                                                  b_val, Lp_val, name=f"{name_prefix}_ShieldInnerVoid", material="vacuum")
            self.hfss.modeler.subtract(shield_outer, [shield_inner_void], keep_originals=False)

            hole_r = b_val + clear_val
            sub_hole = self.hfss.modeler.create_cylinder("Z", [x_feed, y_feed, 0.0],
                                                         hole_r, h_sub_val, name=f"{name_prefix}_SubHole", material="vacuum")
            self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=False)
            g_hole = self.hfss.modeler.create_circle("XY", [x_feed, y_feed, 0.0], hole_r,
                                                     name=f"{name_prefix}_GndHole", material="vacuum")
            self.hfss.modeler.subtract(ground, [g_hole], keep_originals=False)

            port_ring = self.hfss.modeler.create_circle("XY", [x_feed, y_feed, -Lp_val], b_val,
                                                        name=f"{name_prefix}_PortRing", material="vacuum")
            port_hole = self.hfss.modeler.create_circle("XY", [x_feed, y_feed, -Lp_val], a_val,
                                                        name=f"{name_prefix}_PortHole", material="vacuum")
            self.hfss.modeler.subtract(port_ring, [port_hole], keep_originals=False)

            eps_line = min(0.1 * (b_val - a_val), 0.05)
            r_start = a_val + eps_line; r_end = b_val - eps_line
            if r_end <= r_start: r_end = a_val + 0.75 * (b_val - a_val)
            p1 = [x_feed + r_start, y_feed, -Lp_val]; p2 = [x_feed + r_end, y_feed, -Lp_val]

            self.hfss.lumped_port(assignment=port_ring.name, integration_line=[p1, p2],
                                  impedance=50.0, name=f"{name_prefix}_Lumped", renormalize=True)
            if f"{name_prefix}_Lumped" not in self.created_ports:
                self.created_ports.append(f"{name_prefix}_Lumped")
            self.log_message(f"Lumped Port '{name_prefix}_Lumped' created (integration line).")
            return pin, None, shield_outer
        except Exception as e:
            self.log_message(f"Exception in coax creation '{name_prefix}': {e}\nTraceback: {traceback.format_exc()}")
            return None, None, None

    # ---------- Pós-solve helpers ----------
    def _add_post_var(self, name: str, value: str) -> bool:
        try:
            props = [
                "NAME:AllTabs",
                ["NAME:LocalVariableTab",
                 ["NAME:PropServers", "LocalVariables"],
                 ["NAME:NewProps",
                  [f"NAME:{name}", "PropType:=", "PostProcessingVariableProp", "UserDef:=", True, "Value:=", value]]]
            ]
            self.hfss.odesign.ChangeProperty(props)
            self.log_message(f"Post var '{name}' = {value} created.")
            return True
        except Exception as e:
            self.log_message(f"Add post var '{name}' failed: {e}")
            return False

    def _edit_sources_with_vars(self, excitations: List[str], pvars: List[str], phvars: List[str]) -> bool:
        try:
            sol = self.hfss.odesign.GetModule("Solutions")
            header = ["IncludePortPostProcessing:=", False, "SpecifySystemPower:=", False]
            cmd = [header]
            for ex, p, ph in zip(excitations, pvars, phvars):
                cmd.append(["Name:=", ex, "Magnitude:=", p, "Phase:=", ph])
            sol.EditSources(cmd)
            self.log_message(f"Solutions.EditSources applied to {len(excitations)} port(s).")
            return True
        except Exception as e:
            self.log_message(f"EditSources failed: {e}")
            return False

    def _ensure_infinite_sphere(self, name="Infinite Sphere1"):
        try:
            rf = self.hfss.odesign.GetModule("RadField")
            # Try to delete existing with same name to avoid duplicates
            try:
                setups = list(rf.GetSetups())
                if name in setups:
                    rf.DeleteSetup(name)
            except Exception:
                pass
            props = [f"NAME:{name}",
                     "UseCustomRadiationSurface:=", False,
                     "CSDefinition:=", "Theta-Phi",
                     "Polarization:=", "Linear",
                     "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "1deg",
                     "PhiStart:=", "-180deg", "PhiStop:=", "180deg", "PhiStep:=", "1deg",
                     "UseLocalCS:=", False]
            rf.InsertInfiniteSphereSetup(props)
            self.log_message(f"Infinite sphere '{name}' created.")
            return name
        except Exception as e:
            self.log_message(f"Infinite sphere creation failed: {e}")
            return None

    def _postprocess_after_solve(self):
        try:
            exs = self._list_excitations()
            if not exs:
                self.log_message("No excitations found for post-processing.")
                return
            pvars, phvars = [], []
            for i in range(1, len(exs) + 1):
                p = f"p{i}"; ph = f"ph{i}"
                self._add_post_var(p, "1W");  self._add_post_var(ph, "0deg")
                pvars.append(p); phvars.append(ph)
            self._edit_sources_with_vars(exs, pvars, phvars)
            # construir painel de fontes na UI
            self.populate_source_controls(exs)
        except Exception as e:
            self.log_message(f"Postprocess-after-solve error: {e}")

    # ------------- Helpers de solução -------------
    def _fetch_solution(self, expression: str, setup_candidates: Optional[List[str]] = None, **kwargs):
        if setup_candidates is None:
            setup_candidates = ["Setup1 : LastAdaptive", "Setup1:LastAdaptive", "Setup1 : Sweep1", "Setup1:Sweep1"]
        last_err = None
        for setup in setup_candidates:
            try:
                sd = self.hfss.post.get_solution_data(expressions=[expression], setup_sweep_name=setup, **kwargs)
                if sd and hasattr(sd, "primary_sweep_values"):
                    return sd
            except Exception as e:
                last_err = e
        if last_err:
            self.log_message(f"get_solution_data failed for '{expression}': {last_err}")
        return None

    def _shape_series(self, data_obj, npoints: int) -> np.ndarray:
        if isinstance(data_obj, (list, tuple)):
            if len(data_obj) == 0:
                return np.array([], dtype=float)
            if hasattr(data_obj[0], "__len__"):
                arr = np.asarray(data_obj[0], dtype=float)
            else:
                arr = np.asarray(data_obj, dtype=float)
        else:
            arr = np.asarray(data_obj, dtype=float)
        arr = np.atleast_1d(arr).astype(float)
        if npoints > 0 and arr.size not in (1, npoints):
            return np.array([], dtype=float)
        return arr

    def _list_excitations(self) -> List[str]:
        names = []
        try: names = self.hfss.get_excitations_name() or []
        except Exception: names = []
        if not names and self.created_ports:
            names = [f"{p}:1" for p in self.created_ports]
        # ordenar por índice numérico
        def keyfn(s):
            import re
            m = re.search(r"P(\d+)_Lumped", s)
            return int(m.group(1)) if m else 1_000_000
        names.sort(key=keyfn)
        return names

    # ------------- Far Field (cuts & 3D) -------------
    def _get_gain_cut(self, frequency: float, cut: str, fixed_angle_deg: float):
        try:
            expr = "dB(GainTotal)"
            if cut.lower() == "theta":
                variations = {"Freq": f"{frequency}GHz", "Theta": "All", "Phi": f"{fixed_angle_deg}deg"}
                prim = "Theta"
            else:
                variations = {"Freq": f"{frequency}GHz", "Theta": f"{fixed_angle_deg}deg", "Phi": "All"}
                prim = "Phi"
            sd = self._fetch_solution(expr, primary_sweep_variable=prim, variations=variations, context="Infinite Sphere1")
            if not sd or not hasattr(sd, "primary_sweep_values"):
                return None, None
            ang = np.asarray(sd.primary_sweep_values, dtype=float)
            gains = self._shape_series(sd.data_real(), ang.size)
            if gains.size == ang.size and gains.size > 0:
                return ang, gains
            return None, None
        except Exception as e:
            self.log_message(f"Error getting {cut} cut: {e}")
            return None, None

    def _get_gain_3d_grid(self, frequency: float, theta_step=10.0, phi_step=10.0):
        """Varre Phi (fixo) e pega Theta = All para montar grade."""
        try:
            phi_vals = np.arange(-180.0, 180.0 + phi_step, phi_step)
            TH_list, G_list = None, []
            for phi in phi_vals:
                sd = self._fetch_solution(
                    "dB(GainTotal)",
                    primary_sweep_variable="Theta",
                    variations={"Freq": f"{frequency}GHz", "Theta": "All", "Phi": f"{phi}deg"},
                    context="Infinite Sphere1"
                )
                if not sd or not hasattr(sd, "primary_sweep_values"):
                    continue
                th = np.asarray(sd.primary_sweep_values, dtype=float)
                g = self._shape_series(sd.data_real(), th.size)
                if g.size != th.size or g.size == 0:
                    continue
                if TH_list is None:
                    TH_list = th
                else:
                    if th.size != TH_list.size:
                        # pula phi que não bate
                        continue
                G_list.append(g)
            if TH_list is None or len(G_list) == 0:
                return None
            G = np.vstack(G_list).T  # shape (Ntheta, Nphi)
            PH = phi_vals[:G.shape[1]]
            TH = TH_list
            return TH, PH, G
        except Exception as e:
            self.log_message(f"3D grid error: {e}")
            return None

    # ------------- Simulation -------------
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
            self.progress_bar.set(0.25)

            self.hfss.modeler.model_units = "mm"; self.log_message("Model units set to: mm")

            sub_name = self.params["substrate_material"]
            if not self.hfss.materials.checkifmaterialexists(sub_name):
                sub_name = "Custom_Substrate"
                self._ensure_material(sub_name, float(self.params["er"]), float(self.params["tan_d"]))

            L = float(self.calculated_params["patch_length"])
            W = float(self.calculated_params["patch_width"])
            spacing = float(self.calculated_params["spacing"])
            rows = int(self.calculated_params["rows"]); cols = int(self.calculated_params["cols"])
            h_sub = float(self.params["substrate_thickness"])
            sub_w = float(self.calculated_params["substrate_width"]); sub_l = float(self.calculated_params["substrate_length"])

            self._set_design_variables(L, W, spacing, rows, cols, h_sub, sub_w, sub_l)
            self.created_ports.clear()

            self.log_message("Creating substrate")
            substrate = self.hfss.modeler.create_box(["-subW/2", "-subL/2", 0], ["subW", "subL", "h_sub"], "Substrate", sub_name)
            self.log_message("Creating ground plane")
            ground = self.hfss.modeler.create_rectangle("XY", ["-subW/2", "-subL/2", 0], ["subW", "subL"], "Ground", "copper")

            self.log_message(f"Creating {rows*cols} patches in {rows}x{cols} configuration")
            patches = []
            total_w = cols * W + (cols - 1) * spacing; total_l = rows * L + (rows - 1) * spacing
            start_x = -total_w / 2 + W / 2; start_y = -total_l / 2 + L / 2
            self.progress_bar.set(0.35)

            count = 0
            for r in range(rows):
                for c in range(cols):
                    if self.stop_simulation:
                        self.log_message("Simulation stopped by user"); return
                    count += 1; patch_name = f"Patch_{count}"
                    cx = start_x + c * (W + spacing); cy = start_y + r * (L + spacing)
                    origin = [cx - W / 2, cy - L / 2, "h_sub"]
                    self.log_message(f"Creating patch {count} at ({r}, {c})")
                    patch = self.hfss.modeler.create_rectangle("XY", origin, ["patchW", "patchL"], patch_name, "copper")
                    patches.append(patch)

                    if self.params["feed_position"] == "edge": y_feed = cy - 0.5 * L + 0.02 * L
                    else: y_feed = cy - 0.5 * L + 0.30 * L
                    relx = float(self.params["feed_rel_x"]); relx = min(max(relx, 0.0), 1.0)
                    x_feed = cx - 0.5 * W + relx * W

                    pad = self.hfss.modeler.create_circle("XY", [x_feed, y_feed, "h_sub"], "a", f"{patch_name}_Pad", "copper")
                    try: self.hfss.modeler.unite([patch, pad])
                    except Exception: pass

                    self._create_coax_feed_lumped(ground=ground, substrate=substrate, x_feed=x_feed, y_feed=y_feed, name_prefix=f"P{count}")
                    self.progress_bar.set(0.35 + 0.25 * (count / float(rows * cols)))

            if self.stop_simulation:
                self.log_message("Simulation stopped by user"); return

            try:
                names = [ground.name] + [p.name for p in patches]
                self.hfss.assign_perfecte_to_sheets(names)
                self.log_message(f"PerfectE assigned to: {names}")
            except Exception as e:
                self.log_message(f"PerfectE assignment warning: {e}")

            self.log_message("Creating air region + radiation boundary")
            lambda0_mm = self.c / (self.params["sweep_start"] * 1e9) * 1000.0
            pad_mm = float(lambda0_mm) / 4.0
            region = self.hfss.modeler.create_region([pad_mm]*6, is_percentage=False)
            self.hfss.assign_radiation_boundary_to_objects(region)
            self.progress_bar.set(0.65)

            # IMPORTANT: Esfera antes do solve (para ter far-field)
            self._ensure_infinite_sphere("Infinite Sphere1")

            # Setup
            self.log_message("Creating simulation setup")
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02
            # salvar campos de radiação (reforça disponibilidade)
            try:
                setup.props["SaveFields"] = False
                setup.props["SaveRadFields"] = True
            except Exception:
                pass

            self.log_message(f"Creating frequency sweep: {self.params['sweep_type']}")
            stype = self.params["sweep_type"]
            try:
                try:
                    sw = setup.get_sweep("Sweep1");  sw.delete() if sw else None
                except Exception: pass
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
                self.log_message(f"Sweep creation warning: {e}")

            exs = self._list_excitations()
            self.log_message(f"Excitations created: {len(exs)} -> {exs}")

            self.log_message("Validating design")
            try: _ = self.hfss.validate_full_design()
            except Exception as e: self.log_message(f"Validation warning: {e}")

            self.log_message("Starting analysis")
            if self.save_project: self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])
            if self.stop_simulation:
                self.log_message("Simulation stopped by user"); return

            self._postprocess_after_solve()
            self.progress_bar.set(0.9)
            self.log_message("Processing results")
            self.analyze_and_mark_s11()     # também redesenha cortes
            self.refresh_patterns_only()    # inclui 3D
            self.progress_bar.set(1.0)
            self.sim_status_label.configure(text="Simulation completed")
            self.log_message("Simulation completed successfully")
        except Exception as e:
            self.log_message(f"Error in simulation: {e}\nTraceback: {traceback.format_exc()}")
            self.sim_status_label.configure(text=f"Simulation error: {e}")
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    # ------------- S11 / VSWR / |Z| -------------
    def _get_s11_curves(self):
        exs = self._list_excitations()
        if not exs:
            return None
        port_name = exs[0].split(":")[0]

        # tentar por nome
        for expr_tpl in [f"( {port_name},{port_name} )", f"({port_name},{port_name})"]:
            sd_db = self._fetch_solution(f"dB(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1", "Setup1 : LastAdaptive"])
            if sd_db and hasattr(sd_db, "primary_sweep_values"):
                f = np.asarray(sd_db.primary_sweep_values, dtype=float)
                y_db = self._shape_series(sd_db.data_real(), f.size)
                if y_db.size == f.size and f.size > 0:
                    sd_re = self._fetch_solution(f"re(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
                    sd_im = self._fetch_solution(f"im(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
                    reS = self._shape_series(sd_re.data_real(), f.size) if sd_re else None
                    imS = self._shape_series(sd_im.data_real(), f.size) if sd_im else None
                    return f, y_db, reS, imS

        # fallback por índice
        expr_tpl = "(1,1)"
        sd_db = self._fetch_solution(f"dB(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1", "Setup1 : LastAdaptive"])
        if not sd_db:
            return None
        f = np.asarray(sd_db.primary_sweep_values, dtype=float)
        y_db = self._shape_series(sd_db.data_real(), f.size)
        sd_re = self._fetch_solution(f"re(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
        sd_im = self._fetch_solution(f"im(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
        reS = self._shape_series(sd_re.data_real(), f.size) if sd_re else None
        imS = self._shape_series(sd_im.data_real(), f.size) if sd_im else None
        return f, y_db, reS, imS

    def analyze_and_mark_s11(self):
        try:
            # Limpa S11/Imp
            self.ax_s11.clear(); self.ax_imp.clear()

            data = self._get_s11_curves()
            if not data:
                self.log_message("Solution Data failed to load. Check solution, context or expression.")
                self.canvas.draw(); return
            f, s11_db, reS, imS = data
            if f.size == 0 or s11_db.size == 0:
                self.log_message("S11 analysis aborted: empty curve.")
                self.canvas.draw(); return

            # S11 dB
            self.ax_s11.plot(f, s11_db, linewidth=2, label="S11 (dB)")
            self.ax_s11.axhline(y=-10, linestyle='--', alpha=0.7, label='-10 dB')
            self.ax_s11.set_xlabel("Frequency (GHz)")
            self.ax_s11.set_ylabel("S11 (dB)")
            self.ax_s11.set_title("S11 & VSWR")
            self.ax_s11.grid(True, alpha=0.5)

            # VSWR via |S|
            s_abs = 10 ** (s11_db / 20.0)
            s_abs = np.clip(s_abs, 0, 0.999999)
            vswr = (1 + s_abs) / (1 - s_abs)
            ax_v = self.ax_s11.twinx()
            ax_v.plot(f, vswr, linestyle='--', alpha=0.8, label='VSWR')
            ax_v.set_ylabel("VSWR")

            # mínimo
            idx_min = int(np.argmin(s11_db))
            f_res = float(f[idx_min]); s11_min_db = float(s11_db[idx_min])
            self.ax_s11.scatter([f_res], [s11_min_db], s=45, marker="o", zorder=5)
            self.ax_s11.annotate(f"f_res={f_res:.4g} GHz\nS11={s11_min_db:.2f} dB",
                                 (f_res, s11_min_db), textcoords="offset points", xytext=(8, -16))
            cf = float(self.params["frequency"])
            self.ax_s11.axvline(x=cf, linestyle=':', alpha=0.7, color='r', label=f"f0={cf:g} GHz")

            # Impedância |Z|
            Zmag = None; R = X = None
            if reS is not None and imS is not None and reS.size == imS.size == f.size:
                S = reS + 1j*imS
                Z0 = 50.0
                with np.errstate(divide='ignore', invalid='ignore'):
                    Z = Z0 * (1 + S) / (1 - S)
                Zmag = np.abs(Z)
                self.ax_imp.plot(f, Zmag, linewidth=2)
                self.ax_imp.set_xlabel("Frequency (GHz)")
                self.ax_imp.set_ylabel("|Z| (Ω)")
                self.ax_imp.set_title("Input Impedance Magnitude")
                self.ax_imp.grid(True, alpha=0.5)
                Zr = Z[idx_min]
                R = float(np.real(Zr)); X = float(np.imag(Zr))

            # guarda
            self.last_s11_analysis = {"f": f, "s11_db": s11_db, "vswr": vswr,
                                      "Zmag": Zmag, "f_res": f_res,
                                      "R": R, "X": X}

            if R is not None and X is not None:
                self.result_label.configure(text=f"Min @ {f_res:.4g} GHz, S11={s11_min_db:.2f} dB, Z≈{R:.1f} + j{X:.1f} Ω")
            else:
                self.result_label.configure(text=f"Min @ {f_res:.4g} GHz, S11={s11_min_db:.2f} dB")

            # Desenha (mantém cortes atuais)
            self.canvas.draw()
        except Exception as e:
            self.log_message(f"Analyze S11 error: {e}\nTraceback: {traceback.format_exc()}")

    # ------------- Patterns / 3D -------------
    def refresh_patterns_only(self):
        """Atualiza theta/phi e 3D baseando-se na solução já existente."""
        try:
            # limpa e redesenha cortes
            self.ax_th.clear(); self.ax_ph.clear(); self.ax_3d.clear()

            f0 = float(self.params["frequency"])

            th, gth = self._get_gain_cut(f0, cut="theta", fixed_angle_deg=0.0)
            if th is not None and gth is not None:
                self.theta_cut = (th, gth)
                self.ax_th.plot(th, gth, linewidth=2)
                self.ax_th.set_xlabel("Theta (deg)"); self.ax_th.set_ylabel("Gain (dB)")
                self.ax_th.set_title("Radiation Pattern - Theta cut (Phi=0°)")
                self.ax_th.grid(True, alpha=0.5)
            else:
                self.ax_th.text(0.5, 0.5, "Theta-cut gain not available",
                                transform=self.ax_th.transAxes, ha="center", va="center")

            ph, gph = self._get_gain_cut(f0, cut="phi", fixed_angle_deg=90.0)
            if ph is not None and gph is not None:
                self.phi_cut = (ph, gph)
                self.ax_ph.plot(ph, gph, linewidth=2)
                self.ax_ph.set_xlabel("Phi (deg)"); self.ax_ph.set_ylabel("Gain (dB)")
                self.ax_ph.set_title("Radiation Pattern - Phi cut (Theta=90°)")
                self.ax_ph.grid(True, alpha=0.5)
            else:
                self.ax_ph.text(0.5, 0.5, "Phi-cut gain not available",
                                transform=self.ax_ph.transAxes, ha="center", va="center")

            # 3D
            grid = self._get_gain_3d_grid(f0, theta_step=self.params["theta_step"], phi_step=self.params["phi_step"])
            if grid is not None:
                TH_deg, PH_deg, Gdb = grid  # shapes (Nt, Np)
                self.grid3d = grid
                TH = np.deg2rad(TH_deg)[:, None] * np.ones((1, Gdb.shape[1]))
                PH = np.deg2rad(PH_deg)[None, :] * np.ones((Gdb.shape[0], 1))
                # raio proporcional ao ganho linear normalizado
                Glin = 10 ** (Gdb / 20.0)
                Glin = Glin - np.min(Glin)
                if np.max(Glin) > 0: Glin = Glin / np.max(Glin)
                R = 0.2 + 0.8 * Glin
                X = R * np.sin(TH) * np.cos(PH)
                Y = R * np.sin(TH) * np.sin(PH)
                Z = R * np.cos(TH)
                surf = self.ax_3d.plot_surface(X, Y, Z, rstride=1, cstride=1, linewidth=0, antialiased=True,
                                               cmap=cm.jet, shade=True)
                self.ax_3d.set_title("3D Gain Pattern (normalized)")
                self.ax_3d.set_axis_off()
            else:
                self.ax_3d.text2D(0.5, 0.5, "3D pattern not available", transform=self.ax_3d.transAxes, ha="center", va="center")

            self.fig.tight_layout()
            self.canvas.draw()
            self.log_message("Patterns refreshed.")
        except Exception as e:
            self.log_message(f"Refresh patterns error: {e}\nTraceback: {traceback.format_exc()}")

    # ------------- Beamforming UI -------------
    def populate_source_controls(self, excitations: List[str]):
        # limpa frame
        for child in self.src_frame.winfo_children():
            child.destroy()
        self.source_controls.clear()

        if not excitations:
            ctk.CTkLabel(self.src_frame, text="No excitations found.").pack(padx=8, pady=6)
            return

        head = ctk.CTkFrame(self.src_frame); head.pack(fill="x", padx=8, pady=4)
        ctk.CTkLabel(head, text="Beamforming & Refresh", font=ctk.CTkFont(weight="bold")).pack(side="left")
        grid = ctk.CTkFrame(self.src_frame); grid.pack(fill="x", padx=8, pady=6)

        # headers
        ctk.CTkLabel(grid, text="Port", width=120).grid(row=0, column=0, padx=4, pady=4, sticky="w")
        ctk.CTkLabel(grid, text="Power (W)", width=120).grid(row=0, column=1, padx=4, pady=4)
        ctk.CTkLabel(grid, text="Phase (deg)", width=300).grid(row=0, column=2, padx=4, pady=4)

        for i, ex in enumerate(excitations, start=1):
            row = i
            ctk.CTkLabel(grid, text=ex.split(":")[0], width=120).grid(row=row, column=0, padx=4, pady=3, sticky="w")
            p_entry = ctk.CTkEntry(grid, width=100); p_entry.insert(0, "1.0")
            p_entry.grid(row=row, column=1, padx=4, pady=3)
            phase_slider = ctk.CTkSlider(grid, from_=0, to=360, number_of_steps=360)
            phase_slider.set(0); phase_slider.grid(row=row, column=2, padx=6, pady=6, sticky="ew")
            grid.grid_columnconfigure(2, weight=1)
            self.source_controls[ex] = {"power": p_entry, "phase": phase_slider}

    def apply_sources_from_ui(self):
        try:
            exs = self._list_excitations()
            if not exs:
                self.log_message("No excitations to apply.")
                return
            pvars, phvars = [], []
            for i, ex in enumerate(exs, start=1):
                ctrl = self.source_controls.get(ex)
                if not ctrl:
                    continue
                try:
                    pw = float(ctrl["power"].get())
                except Exception:
                    pw = 1.0
                ph = float(ctrl["phase"].get())
                self._add_post_var(f"p{i}", f"{pw}W")
                self._add_post_var(f"ph{i}", f"{ph}deg")
                pvars.append(f"p{i}"); phvars.append(f"ph{i}")
            self._edit_sources_with_vars(exs, pvars, phvars)
            self.refresh_patterns_only()
        except Exception as e:
            self.log_message(f"Apply sources error: {e}\nTraceback: {traceback.format_exc()}")

    def toggle_auto_refresh(self):
        if self.auto_refresh_var.get():
            self.schedule_auto_refresh()
        else:
            if self.auto_refresh_job:
                try: self.window.after_cancel(self.auto_refresh_job)
                except Exception: pass
                self.auto_refresh_job = None

    def schedule_auto_refresh(self):
        self.refresh_patterns_only()
        if self.auto_refresh_var.get():
            self.auto_refresh_job = self.window.after(1500, self.schedule_auto_refresh)

    # ------------- Export -------------
    def export_csv(self):
        try:
            if self.last_s11_analysis:
                f = self.last_s11_analysis["f"]; s11 = self.last_s11_analysis["s11_db"]
                np.savetxt("simulation_results.csv", np.column_stack((f, s11)),
                           delimiter=",", header="Frequency (GHz), S11 (dB)", comments='')
                self.log_message("Data exported to simulation_results.csv")
            else:
                self.log_message("Run 'Analyze S11' first.")
        except Exception as e:
            self.log_message(f"Error exporting CSV: {e}")

    def export_png(self):
        try:
            if hasattr(self, 'fig'):
                self.fig.savefig("simulation_results.png", dpi=300, bbox_inches='tight')
                self.log_message("Plot saved to simulation_results.png")
        except Exception as e:
            self.log_message(f"Error saving plot: {e}")

    # ------------- Cleanup -------------
    def cleanup(self):
        try:
            if self.hfss:
                try:
                    if self.save_project: self.hfss.save_project()
                    else: self.hfss.close_project(save=False)
                except Exception as e:
                    self.log_message(f"Error closing project: {e}")
            if self.desktop:
                try: self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception as e: self.log_message(f"Error releasing desktop: {e}")
            if self.temp_folder and not self.save_project:
                try: self.temp_folder.cleanup()
                except Exception as e: self.log_message(f"Error cleaning temp files: {e}")
        except Exception as e:
            self.log_message(f"Error during cleanup: {e}")

    def on_closing(self):
        self.log_message("Application closing...")
        self.cleanup()
        self.window.quit(); self.window.destroy()

    # ------------- Persistence -------------
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
                if k in all_params: self.params[k] = all_params[k]
            for k in self.calculated_params:
                if k in all_params: self.calculated_params[k] = all_params[k]
            self.update_interface_from_params()
            self.log_message("Parameters loaded from antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error loading parameters: {e}")

    def update_interface_from_params(self):
        try:
            for key, widget in self.entries:
                if key in self.params:
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end"); widget.insert(0, str(self.params[key]))
                    elif isinstance(widget, ctk.StringVar): widget.set(self.params[key])
                    elif isinstance(widget, ctk.BooleanVar): widget.set(self.params[key])
            self.patches_label.configure(text=f"Number of Patches: {self.calculated_params['num_patches']}")
            self.rows_cols_label.configure(text=f"Configuration: {self.calculated_params['rows']} x {self.calculated_params['cols']}")
            self.spacing_label.configure(text=f"Spacing: {self.calculated_params['spacing']:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch Dimensions: {self.calculated_params['patch_length']:.2f} x "
                                                 f"{self.calculated_params['patch_width']:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {self.calculated_params['lambda_g']:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset (y): {self.calculated_params['feed_offset']:.2f} mm")
            self.substrate_dims_label.configure(text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x "
                                                     f"{self.calculated_params['substrate_length']:.2f} mm")
            self.log_message("Interface updated with loaded parameters")
        except Exception as e:
            self.log_message(f"Error updating interface: {e}")

    def run(self):
        try:
            self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.window.mainloop()
        finally:
            self.cleanup()


if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()

    app.run()
