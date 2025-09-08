# -*- coding: utf-8 -*-
"""
Modern Patch Antenna Designer (v2.2)
------------------------------------
Novidades nesta versão:
- Opção de stack-up completo com modo 'Aperture-Coupled (2 slots + dipoles)'.
- Dois dipolos na camada inferior alimentando duas fendas ortogonais no GND intermediário.
- Camada opcional 'Honeycomb' (homogeneizada) acima do patch.
- Correções nas rotinas de padrões 2D/3D e em EditSources.
- Correção de indentação e fluxo de execução/thread.
- Mantido modo 'Array Coaxial' do projeto original.

Requer: Ansys Electronics Desktop + PyAEDT compatível.
"""

import os
import re
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
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401
import customtkinter as ctk

from ansys.aedt.core import Desktop, Hfss

# ---------- Aparência ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ModernPatchAntennaDesigner:
    """GUI para dimensionamento e simulação de antena patch/array em HFSS."""

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

        # Dados em memória
        self.last_s11_analysis = None
        self.theta_cut = None
        self.phi_cut = None
        self.grid3d = None
        self.auto_refresh_job = None

        # Parâmetros do usuário (defaults)
        self.params = {
            # RF básicos
            "frequency": 10.0,             # GHz
            "gain": 12.0,                  # dBi (apenas para dimensionar array coaxial)
            "sweep_start": 8.0,            # GHz
            "sweep_stop": 12.0,            # GHz
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,
            "sweep_type": "Interpolating",  # Discrete | Interpolating | Fast
            "sweep_step": 0.02,            # GHz (Discrete)

            # Tipo de arquitetura
            "feed_scheme": "Array Coaxial",  # Array Coaxial | Aperture-Coupled (2 slots + dipoles)

            # Array (somente para coaxial)
            "spacing_type": "lambda/2",

            # Substrato superior (onde fica o patch no modo coaxial; no modo aperture é o espaçador)
            "substrate_material": "Duroid (tm)",
            "substrate_thickness": 0.5,    # mm (coaxial) — No aperture, torna-se 'spacer_thk'
            "metal_thickness": 0.035,      # mm
            "er": 2.2,
            "tan_d": 0.0009,

            # Alimentação coaxial (modo coaxial)
            "feed_position": "inset",      # edge|inset
            "feed_rel_x": 0.485,
            "probe_radius": 0.40,          # mm
            "coax_ba_ratio": 2.3,
            "coax_wall_thickness": 0.20,   # mm
            "coax_port_length": 3.0,       # mm
            "antipad_clearance": 0.10,     # mm

            # Amostragem 3D
            "theta_step": 10.0,            # deg (p/ grid de pós-processamento)
            "phi_step": 10.0,              # deg

            # ---------- Parâmetros modo Aperture-Coupled ----------
            # Substrato de alimentação (abaixo do GND com fendas)
            "feed_sub_thk": 0.8,           # mm
            "feed_sub_er": 2.94,           # RO4003C típico
            "feed_sub_tand": 0.0027,

            # GND intermediário com duas fendas
            "slot_len": 9.0,               # mm (ajuste fino ~ 0.8..1.0*L_patch)
            "slot_w": 0.8,                 # mm
            "slot_shift": 0.0,             # mm (deslocamento do centro, se desejar)

            # Espaçador entre GND fendido e patch (espuma/ar)
            "spacer_thk": 3.0,             # mm
            "spacer_er": 1.05,
            "spacer_tand": 0.0001,

            # Dipolos impressos na face inferior do GND (na face superior do feed_sub)
            "dipole_len": 9.0,             # mm (por braço)
            "dipole_w": 1.2,               # mm
            "dipole_gap": 0.4,             # mm (folga central p/ porta)
            "port_impedance": 50.0,

            # Patch superior (mesmo cálculo L/W do coaxial)
            # (usamos er do 'spacer' para eeff no cálculo aproximado do patch)
            # Camada opcional Honeycomb acima do patch
            "include_honeycomb": False,
            "hc_thk": 3.0,                 # mm
            "hc_er": 1.2,
            "hc_tand": 0.0001,
        }

        # Parâmetros calculados
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

        def add_entry(section, label, key, value, row, combo=None, check=False, width=220):
            ctk.CTkLabel(section, text=label, font=ctk.CTkFont(weight="bold")
                         ).grid(row=row, column=0, padx=15, pady=6, sticky="w")
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(section, values=combo, variable=var, width=width)
                widget.grid(row=row, column=1, padx=15, pady=6)
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(section, text="", variable=var)
                widget.grid(row=row, column=1, padx=15, pady=6, sticky="w")
                self.entries.append((key, var))
            else:
                widget = ctk.CTkEntry(section, width=width)
                widget.insert(0, str(value))
                widget.grid(row=row, column=1, padx=15, pady=6)
                self.entries.append((key, widget))
            return row + 1

        row_idx = add_entry(sec_ant, "Central Frequency (GHz):", "frequency", self.params["frequency"], row_idx)
        row_idx = add_entry(sec_ant, "Desired Gain (dBi) [coax array]:", "gain", self.params["gain"], row_idx)
        row_idx = add_entry(sec_ant, "Sweep Start (GHz):", "sweep_start", self.params["sweep_start"], row_idx)
        row_idx = add_entry(sec_ant, "Sweep Stop (GHz):", "sweep_stop", self.params["sweep_stop"], row_idx)
        row_idx = add_entry(sec_ant, "Feed/Stack Scheme:", "feed_scheme", self.params["feed_scheme"], row_idx,
                            combo=["Array Coaxial", "Aperture-Coupled (2 slots + dipoles)"], width=280)
        row_idx = add_entry(sec_ant, "Patch Spacing [coax array]:", "spacing_type", self.params["spacing_type"], row_idx,
                            combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        # Substrato/coax (modo coaxial) e também usa 'metal_thickness'
        sec_sub = self.create_section(main, "Upper Substrate / Metal", 1, 0)
        row_idx = 2
        row_idx = add_entry(sec_sub, "Substrate Material:", "substrate_material",
                            self.params["substrate_material"], row_idx,
                            combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        row_idx = add_entry(sec_sub, "Relative Permittivity (εr):", "er", self.params["er"], row_idx)
        row_idx = add_entry(sec_sub, "Loss Tangent (tan δ):", "tan_d", self.params["tan_d"], row_idx)
        row_idx = add_entry(sec_sub, "Substrate Thickness (mm) [coax] / Spacer (mm) [aperture]:",
                            "substrate_thickness", self.params["substrate_thickness"], row_idx)
        row_idx = add_entry(sec_sub, "Metal Thickness (mm):", "metal_thickness", self.params["metal_thickness"], row_idx)

        # Coax
        sec_coax = self.create_section(main, "Coaxial Feed [Array Coaxial]", 2, 0)
        row_idx = 2
        row_idx = add_entry(sec_coax, "Feed position type:", "feed_position", self.params["feed_position"], row_idx,
                            combo=["inset", "edge"])
        row_idx = add_entry(sec_coax, "Feed relative X (0..1):", "feed_rel_x", self.params["feed_rel_x"], row_idx)
        row_idx = add_entry(sec_coax, "Inner radius a (mm):", "probe_radius", self.params["probe_radius"], row_idx)
        row_idx = add_entry(sec_coax, "b/a ratio:", "coax_ba_ratio", self.params["coax_ba_ratio"], row_idx)
        row_idx = add_entry(sec_coax, "Shield wall (mm):", "coax_wall_thickness", self.params["coax_wall_thickness"], row_idx)
        row_idx = add_entry(sec_coax, "Port length below GND Lp (mm):", "coax_port_length", self.params["coax_port_length"], row_idx)
        row_idx = add_entry(sec_coax, "Anti-pad clearance (mm):", "antipad_clearance", self.params["antipad_clearance"], row_idx)

        # Aperture-coupled
        sec_ap = self.create_section(main, "Aperture-Coupled Parameters", 3, 0)
        row_idx = 2
        row_idx = add_entry(sec_ap, "Feed substrate thickness (mm):", "feed_sub_thk", self.params["feed_sub_thk"], row_idx)
        row_idx = add_entry(sec_ap, "Feed substrate εr:", "feed_sub_er", self.params["feed_sub_er"], row_idx)
        row_idx = add_entry(sec_ap, "Feed substrate tanδ:", "feed_sub_tand", self.params["feed_sub_tand"], row_idx)
        row_idx = add_entry(sec_ap, "Slot length (mm):", "slot_len", self.params["slot_len"], row_idx)
        row_idx = add_entry(sec_ap, "Slot width (mm):", "slot_w", self.params["slot_w"], row_idx)
        row_idx = add_entry(sec_ap, "Slot center shift (mm):", "slot_shift", self.params["slot_shift"], row_idx)
        row_idx = add_entry(sec_ap, "Spacer thickness (mm):", "spacer_thk", self.params["spacer_thk"], row_idx)
        row_idx = add_entry(sec_ap, "Spacer εr:", "spacer_er", self.params["spacer_er"], row_idx)
        row_idx = add_entry(sec_ap, "Spacer tanδ:", "spacer_tand", self.params["spacer_tand"], row_idx)
        row_idx = add_entry(sec_ap, "Dipole arm length (mm):", "dipole_len", self.params["dipole_len"], row_idx)
        row_idx = add_entry(sec_ap, "Dipole width (mm):", "dipole_w", self.params["dipole_w"], row_idx)
        row_idx = add_entry(sec_ap, "Dipole gap (mm):", "dipole_gap", self.params["dipole_gap"], row_idx)
        row_idx = add_entry(sec_ap, "Port impedance (Ω):", "port_impedance", self.params["port_impedance"], row_idx)
        row_idx = add_entry(sec_ap, "Include Honeycomb Layer Above Patch:", "include_honeycomb", self.params["include_honeycomb"], row_idx, check=True)
        row_idx = add_entry(sec_ap, "Honeycomb thickness (mm):", "hc_thk", self.params["hc_thk"], row_idx)
        row_idx = add_entry(sec_ap, "Honeycomb εr eff:", "hc_er", self.params["hc_er"], row_idx)
        row_idx = add_entry(sec_ap, "Honeycomb tanδ:", "hc_tand", self.params["hc_tand"], row_idx)

        # Simulação
        sec_sim = self.create_section(main, "Simulation Settings", 4, 0)
        row_idx = 2
        row_idx = add_entry(sec_sim, "CPU Cores:", "cores", self.params["cores"], row_idx)
        row_idx = add_entry(sec_sim, "Show HFSS Interface:", "show_gui", not self.params["non_graphical"], row_idx, check=True)
        row_idx = add_entry(sec_sim, "Save Project:", "save_project", self.save_project, row_idx, check=True)
        row_idx = add_entry(sec_sim, "Sweep Type:", "sweep_type", self.params["sweep_type"], row_idx,
                            combo=["Discrete", "Interpolating", "Fast"])
        row_idx = add_entry(sec_sim, "Discrete Step (GHz):", "sweep_step", self.params["sweep_step"], row_idx)
        row_idx = add_entry(sec_sim, "3D Theta step (deg):", "theta_step", self.params["theta_step"], row_idx)
        row_idx = add_entry(sec_sim, "3D Phi step (deg):", "phi_step", self.params["phi_step"], row_idx)

        # Calculados
        sec_calc = self.create_section(main, "Calculated Parameters", 5, 0)
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
        ctk.CTkLabel(note, text="Tip: With direct sources (Magnitude/Phase) you can retune beams without re-solving.",
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

    # ---------- Log utils ----------
    def log_message(self, message: str):
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

    # ----------- Física / Cálculos -----------
    def _validate_ranges(self) -> bool:
        ok = True
        msgs = []
        if self.params["frequency"] <= 0:
            ok = False; msgs.append("frequency must be > 0")
        if self.params["sweep_start"] <= 0 or self.params["sweep_stop"] <= 0:
            ok = False; msgs.append("sweep_start/stop must be > 0")
        if self.params["sweep_start"] >= self.params["sweep_stop"]:
            ok = False; msgs.append("sweep_start must be < sweep_stop")
        if self.params["er"] < 1:
            ok = False; msgs.append("er must be >= 1")
        if self.params["substrate_thickness"] <= 0:
            ok = False; msgs.append("substrate_thickness/spacer_thk must be > 0")
        if not (0.0 <= self.params["feed_rel_x"] <= 1.0):
            ok = False; msgs.append("feed_rel_x must be in [0,1]")
        if self.params["probe_radius"] <= 0:
            ok = False; msgs.append("probe_radius must be > 0")
        if self.params["coax_ba_ratio"] <= 1.05:
            ok = False; msgs.append("coax_ba_ratio must be > 1.05")
        if self.params["coax_port_length"] <= 0:
            ok = False; msgs.append("coax_port_length must be > 0")
        if self.params["theta_step"] <= 0 or self.params["phi_step"] <= 0:
            ok = False; msgs.append("theta_step/phi_step must be > 0")
        if not ok:
            msg = "; ".join(msgs)
            self.status_label.configure(text=f"Invalid parameters: {msg}")
            self.log_message(f"Invalid parameters: {msg}")
        return ok

    def get_parameters(self) -> bool:
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
                             "sweep_step", "theta_step", "phi_step",
                             "feed_sub_thk", "feed_sub_er", "feed_sub_tand",
                             "slot_len", "slot_w", "slot_shift",
                             "spacer_thk", "spacer_er", "spacer_tand",
                             "dipole_len", "dipole_w", "dipole_gap", "port_impedance",
                             "hc_thk", "hc_er", "hc_tand"]:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
                elif key in ["spacing_type", "substrate_material", "feed_position", "sweep_type", "feed_scheme"]:
                    self.params[key] = widget.get()
                elif key in ["include_honeycomb"]:
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
        return self._validate_ranges()

    def calculate_patch_dimensions(self, frequency_ghz: float, er_eff: Optional[float] = None) -> Tuple[float, float, float]:
        """Calcula L, W e λg (mm) para retangular. Se er_eff for None, usa self.params['er']."""
        f = frequency_ghz * 1e9
        er_local = float(self.params["er"]) if er_eff is None else float(er_eff)
        h = float(self.params["substrate_thickness"]) / 1000.0  # mm->m (no modo aperture, este valor não é usado para patch)
        W = self.c / (2 * f) * math.sqrt(2 / (er_local + 1))
        eeff = (er_local + 1) / 2 + (er_local - 1) / 2 * (1 + 12 * h / W) ** -0.5
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        L_eff = self.c / (2 * f * math.sqrt(eeff))
        L = L_eff - 2 * dL
        lambda_g = self.c / (f * math.sqrt(eeff))
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)

    def _size_array_from_gain(self) -> Tuple[int, int, int]:
        G_elem = 8.0
        G_des = float(self.params["gain"])
        N_req = max(1, int(math.ceil(10 ** ((G_des - G_elem) / 10.0))))
        if N_req % 2 == 1:
            N_req += 1
        rows = max(2, int(round(math.sqrt(N_req))));  rows += rows % 2
        cols = max(2, int(math.ceil(N_req / rows))); cols += cols % 2
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
            # Para o modo aperture, usamos er ~ do espaçador para calcular patch
            if self.params["feed_scheme"].startswith("Aperture"):
                L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"],
                                                                          er_eff=self.params["spacer_er"])
                rows, cols = 1, 1  # força 1x1
            else:
                L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"])
                rows, cols, _ = self._size_array_from_gain()

            self.calculated_params.update({"patch_length": L_mm, "patch_width": W_mm, "lambda_g": lambda_g_mm})
            lambda0_m = self.c / (self.params["frequency"] * 1e9)
            factors = {"lambda/2": 0.5, "lambda": 1.0, "0.7*lambda": 0.7, "0.8*lambda": 0.8, "0.9*lambda": 0.9}
            spacing_mm = factors.get(self.params["spacing_type"], 0.5) * lambda0_m * 1000.0
            self.calculated_params["spacing"] = spacing_mm
            self.calculated_params.update({"num_patches": rows * cols, "rows": rows, "cols": cols})
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

    def _set_design_variables_common(self):
        self.hfss["f0"] = f"{self.params['frequency']}GHz"
        self.hfss["t_met"] = f"{self.params['metal_thickness']}mm"
        self.hfss["eps"] = "0.001mm"

    def _ensure_infinite_sphere(self, name="Infinite Sphere1") -> Optional[str]:
        try:
            rf = self.hfss.odesign.GetModule("RadField")
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
                     "ThetaStart:=", "0deg", "ThetaStop:=", "180deg", "ThetaStep:=", "1deg",
                     "PhiStart:=", "-180deg", "PhiStop:=", "180deg", "PhiStep:=", "1deg",
                     "UseLocalCS:=", False]
            rf.InsertInfiniteSphereSetup(props)
            self.log_message(f"Infinite sphere '{name}' created.")
            return name
        except Exception as e:
            self.log_message(f"Infinite sphere creation failed: {e}")
            return None

    # ---------- Geometrias ----------
    def _build_array_coax(self):
        """Constrói o array como no projeto original (coaxial)."""
        L = float(self.calculated_params["patch_length"])
        W = float(self.calculated_params["patch_width"])
        spacing = float(self.calculated_params["spacing"])
        rows = int(self.calculated_params["rows"]); cols = int(self.calculated_params["cols"])
        h_sub = float(self.params["substrate_thickness"])
        sub_w = float(self.calculated_params["substrate_width"]); sub_l = float(self.calculated_params["substrate_length"])

        # Material
        sub_name = self.params["substrate_material"]
        if not self.hfss.materials.checkifmaterialexists(sub_name):
            sub_name = "Custom_Substrate"
            self._ensure_material(sub_name, float(self.params["er"]), float(self.params["tan_d"]))

        self.hfss["patchL"] = f"{L}mm"; self.hfss["patchW"] = f"{W}mm"
        self.hfss["spacing"] = f"{spacing}mm"; self.hfss["rows"] = str(rows); self.hfss["cols"] = str(cols)
        self.hfss["h_sub"] = f"{h_sub}mm"
        self.hfss["subW"] = f"{sub_w}mm"; self.hfss["subL"] = f"{sub_l}mm"
        a = float(self.params["probe_radius"]); b = a * float(self.params["coax_ba_ratio"])
        wall = float(self.params["coax_wall_thickness"]); Lp = float(self.params["coax_port_length"])
        clear = float(self.params["antipad_clearance"])
        self.hfss["a"] = f"{a}mm"; self.hfss["b"] = f"{b}mm"; self.hfss["wall"] = f"{wall}mm"
        self.hfss["Lp"] = f"{Lp}mm"; self.hfss["clear"] = f"{clear}mm"
        self.hfss["padAir"] = f"{max(spacing, W, L)/2 + Lp + 2.0}mm"

        self.log_message("Creating substrate (coaxial)")
        substrate = self.hfss.modeler.create_box(
            ["-subW/2", "-subL/2", 0], ["subW", "subL", "h_sub"], name="Substrate", matname=sub_name
        )
        self.log_message("Creating ground plane")
        ground = self.hfss.modeler.create_rectangle(
            "XY", ["-subW/2", "-subL/2", 0], ["subW", "subL"], name="Ground", matname="copper"
        )

        self.log_message(f"Creating {rows*cols} patches in {rows}x{cols} configuration")
        patches = []
        total_w = cols * W + (cols - 1) * spacing; total_l = rows * L + (rows - 1) * spacing
        start_x = -total_w / 2 + W / 2; start_y = -total_l / 2 + L / 2

        count = 0
        for r in range(rows):
            for c in range(cols):
                count += 1; patch_name = f"Patch_{count}"
                cx = start_x + c * (W + spacing); cy = start_y + r * (L + spacing)
                origin = [cx - W / 2, cy - L / 2, "h_sub"]
                patch = self.hfss.modeler.create_rectangle(
                    "XY", origin, ["patchW", "patchL"], name=patch_name, matname="copper"
                )
                patches.append(patch)

                if self.params["feed_position"] == "edge":
                    y_feed = cy - 0.5 * L + 0.02 * L
                else:
                    y_feed = cy - 0.5 * L + 0.30 * L
                relx = float(self.params["feed_rel_x"]); relx = min(max(relx, 0.0), 1.0)
                x_feed = cx - 0.5 * W + relx * W
                pad = self.hfss.modeler.create_circle(
                    "XY", [x_feed, y_feed, "h_sub"], "a", name=f"{patch_name}_Pad", matname="copper"
                )
                try:
                    self.hfss.modeler.unite([patch, pad])
                except Exception:
                    pass
                self._create_coax_feed_lumped(ground, substrate, x_feed, y_feed, f"P{count}")

        try:
            names = [ground.name] + [p.name for p in patches]
            self.hfss.assign_perfecte_to_sheets(names)
            self.log_message(f"PerfectE assigned to: {names}")
        except Exception as e:
            self.log_message(f"PerfectE assignment warning: {e}")

        return ground, patches

    def _build_aperture_coupled_stack(self):
        """Single element: dipolos (2 portas) -> fendas no GND -> patch -> (opcional) honeycomb."""
        # Nomes de materiais
        feed_mat = "Feed_Substrate"
        spacer_mat = "Spacer_Mat"
        if not self.hfss.materials.checkifmaterialexists(feed_mat):
            self._ensure_material(feed_mat, self.params["feed_sub_er"], self.params["feed_sub_tand"])
        if not self.hfss.materials.checkifmaterialexists(spacer_mat):
            self._ensure_material(spacer_mat, self.params["spacer_er"], self.params["spacer_tand"])
        if self.params["include_honeycomb"] and not self.hfss.materials.checkifmaterialexists("Honeycomb_Eff"):
            self._ensure_material("Honeycomb_Eff", self.params["hc_er"], self.params["hc_tand"])

        # Dimensões do patch (já calculadas com er~spacer)
        Lp = float(self.calculated_params["patch_length"])
        Wp = float(self.calculated_params["patch_width"])
        # Dimensão do plano de referência (região útil)
        slab_w = max(6*Wp, 6*Lp)  # grande o bastante para região local
        slab_l = slab_w

        # Variáveis p/ HFSS
        self.hfss["patchL"] = f"{Lp}mm"; self.hfss["patchW"] = f"{Wp}mm"
        self.hfss["slabW"] = f"{slab_w}mm"; self.hfss["slabL"] = f"{slab_l}mm"
        self.hfss["t_feed"] = f"{self.params['feed_sub_thk']}mm"
        self.hfss["t_spacer"] = f"{self.params['spacer_thk']}mm"
        self.hfss["slotL"] = f"{self.params['slot_len']}mm"
        self.hfss["slotW"] = f"{self.params['slot_w']}mm"
        self.hfss["slotShift"] = f"{self.params['slot_shift']}mm"

        z0 = 0.0
        # 1) Substrato de alimentação
        feed_sub = self.hfss.modeler.create_box(
            ["-slabW/2", "-slabL/2", z0], ["slabW", "slabL", "t_feed"],
            name="FeedSub", matname=feed_mat
        )
        # 2) GND intermediário (com slots) na face superior do feed_sub
        gnd_mid = self.hfss.modeler.create_rectangle(
            "XY", ["-slabW/2", "-slabL/2", "t_feed"], ["slabW", "slabL"],
            name="GND_MID", matname="copper"
        )
        # Slots: dois retângulos ortogonais e centrados
        # Slot-X
        sx = self.hfss.modeler.create_rectangle(
            "XY", ["-slotL/2+slotShift", "-slotW/2", "t_feed"], ["slotL", "slotW"],
            name="SlotX", matname="vacuum"
        )
        # Slot-Y
        sy = self.hfss.modeler.create_rectangle(
            "XY", ["-slotW/2", "-slotL/2+slotShift", "t_feed"], ["slotW", "slotL"],
            name="SlotY", matname="vacuum"
        )
        try:
            self.hfss.modeler.subtract(gnd_mid, [sx, sy], keep_originals=False)
            self.log_message("Created orthogonal slots in middle ground.")
        except Exception as e:
            self.log_message(f"Slot subtraction warning: {e}")

        # 3) Espaçador entre GND e Patch
        spacer = self.hfss.modeler.create_box(
            ["-slabW/2", "-slabL/2", "t_feed"], ["slabW", "slabL", "t_spacer"],
            name="Spacer", matname=spacer_mat
        )

        # 4) Patch no topo do espaçador
        patch = self.hfss.modeler.create_rectangle(
            "XY", ["-patchW/2", "-patchL/2", "t_feed+t_spacer"], ["patchW", "patchL"],
            name="Patch_TOP", matname="copper"
        )

        # 5) Dipolos impressos (duas portas) sob o GND (no topo do feed_sub)
        z_dip = float(self.params["feed_sub_thk"]) - 0.001  # ligeiramente abaixo para não tocar o GND
        dl = float(self.params["dipole_len"])
        dw = float(self.params["dipole_w"])
        gap = float(self.params["dipole_gap"])
        # Dipolo X (alinha com SlotX)
        # braços esquerda/direita ao longo de x
        dx1 = self.hfss.modeler.create_rectangle(
            "XY", [-gap/2 - dl, -dw/2, z_dip], [dl, dw], name="DipX_Left", matname="copper"
        )
        dx2 = self.hfss.modeler.create_rectangle(
            "XY", [gap/2, -dw/2, z_dip], [dl, dw], name="DipX_Right", matname="copper"
        )
        # Porta lumped (folha no gap)
        px = self.hfss.modeler.create_rectangle(
            "XY", [-gap/2, -dw/2, z_dip], [gap, dw], name="PortSheet_X", matname="vacuum"
        )

        # Dipolo Y (alinha com SlotY): braços inferior/superior ao longo de y
        dy1 = self.hfss.modeler.create_rectangle(
            "XY", [-dw/2, -gap/2 - dl, z_dip], [dw, dl], name="DipY_Bottom", matname="copper"
        )
        dy2 = self.hfss.modeler.create_rectangle(
            "XY", [-dw/2, gap/2, z_dip], [dw, dl], name="DipY_Top", matname="copper"
        )
        py = self.hfss.modeler.create_rectangle(
            "XY", [-dw/2, -gap/2, z_dip], [dw, gap], name="PortSheet_Y", matname="vacuum"
        )

        # Atribuir PEC aos metais de folha
        try:
            self.hfss.assign_perfecte_to_sheets([gnd_mid.name, patch.name,
                                                 dx1.name, dx2.name, dy1.name, dy2.name])
        except Exception as e:
            self.log_message(f"PerfectE assignment warning (aperture): {e}")

        # 6) Honeycomb opcional acima do patch
        z_top = float(self.params["feed_sub_thk"] + self.params["spacer_thk"])
        honey = None
        if self.params["include_honeycomb"]:
            self.hfss["t_hc"] = f"{self.params['hc_thk']}mm"
            honey = self.hfss.modeler.create_box(
                ["-slabW/2", "-slabL/2", z_top], ["slabW", "slabL", "t_hc"],
                name="HoneycombLayer", matname="Honeycomb_Eff"
            )

        # 7) Portas lumped nos gaps dos dipolos
        zport = z_dip
        # Porta X - integração de baixo para cima no retângulo do gap
        p1x = [0.0, -dw/2, zport]; p2x = [0.0, +dw/2, zport]
        self.hfss.lumped_port(assignment=px.name, integration_line=[p1x, p2x],
                              impedance=self.params["port_impedance"], name="Port_X", renormalize=True)
        # Porta Y
        p1y = [-dw/2, 0.0, zport]; p2y = [+dw/2, 0.0, zport]
        self.hfss.lumped_port(assignment=py.name, integration_line=[p1y, p2y],
                              impedance=self.params["port_impedance"], name="Port_Y", renormalize=True)
        self.created_ports = ["Port_X", "Port_Y"]

        # Retorno: objetos principais
        return dict(feed_sub=feed_sub, gnd_mid=gnd_mid, spacer=spacer, patch=patch, honey=honey)

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
                "Z", [x_feed, y_feed, -Lp_val], a_val, h_sub_val + Lp_val + 0.001,
                name=f"{name_prefix}_Pin", matname="copper"
            )
            shield_outer = self.hfss.modeler.create_cylinder(
                "Z", [x_feed, y_feed, -Lp_val], b_val + wall_val, Lp_val,
                name=f"{name_prefix}_ShieldOuter", matname="copper"
            )
            shield_inner_void = self.hfss.modeler.create_cylinder(
                "Z", [x_feed, y_feed, -Lp_val], b_val, Lp_val,
                name=f"{name_prefix}_ShieldInnerVoid", matname="vacuum"
            )
            self.hfss.modeler.subtract(shield_outer, [shield_inner_void], keep_originals=False)

            hole_r = b_val + clear_val
            sub_hole = self.hfss.modeler.create_cylinder(
                "Z", [x_feed, y_feed, 0.0], hole_r, h_sub_val,
                name=f"{name_prefix}_SubHole", matname="vacuum"
            )
            self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=False)
            g_hole = self.hfss.modeler.create_circle(
                "XY", [x_feed, y_feed, 0.0], hole_r,
                name=f"{name_prefix}_GndHole", matname="vacuum"
            )
            self.hfss.modeler.subtract(ground, [g_hole], keep_originals=False)

            port_ring = self.hfss.modeler.create_circle(
                "XY", [x_feed, y_feed, -Lp_val], b_val,
                name=f"{name_prefix}_PortRing", matname="vacuum"
            )
            port_hole = self.hfss.modeler.create_circle(
                "XY", [x_feed, y_feed, -Lp_val], a_val,
                name=f"{name_prefix}_PortHole", matname="vacuum"
            )
            self.hfss.modeler.subtract(port_ring, [port_hole], keep_originals=False)

            eps_line = min(0.1 * (b_val - a_val), 0.05)
            r_start = a_val + eps_line; r_end = b_val - eps_line
            if r_end <= r_start:
                r_end = a_val + 0.75 * (b_val - a_val)
            p1 = [x_feed + r_start, y_feed, -Lp_val]; p2 = [x_feed + r_end, y_feed, -Lp_val]

            self.hfss.lumped_port(
                assignment=port_ring.name, integration_line=[p1, p2],
                impedance=50.0, name=f"{name_prefix}_Lumped", renormalize=True
            )
            if f"{name_prefix}_Lumped" not in self.created_ports:
                self.created_ports.append(f"{name_prefix}_Lumped")
            self.log_message(f"Lumped Port '{name_prefix}_Lumped' created (integration line).")
            return pin, None, shield_outer
        except Exception as e:
            self.log_message(f"Exception in coax creation '{name_prefix}': {e}\nTraceback: {traceback.format_exc()}")
            return None, None, None

    # ---------- Pós-solve helpers ----------
    def _edit_sources_with_vars(self, excitations: List[str], magnitudes: List[str], phases: List[str]) -> bool:
        try:
            sol = self.hfss.odesign.GetModule("Solutions")
            header = ["IncludePortPostProcessing:=", False, "SpecifySystemPower:=", False]
            cmd = [header]
            for ex, mag, ph in zip(excitations, magnitudes, phases):
                cmd.append(["Name:=", ex, "Magnitude:=", mag, "Phase:=", ph])
            sol.EditSources(cmd)
            self.log_message(f"Solutions.EditSources applied to {len(excitations)} port(s).")
            return True
        except Exception as e:
            self.log_message(f"EditSources failed: {e}")
            return False

    def _postprocess_after_solve(self):
        try:
            exs = self._list_excitations()
            if not exs:
                self.log_message("No excitations found for post-processing.")
                return
            self._edit_sources_with_vars(exs, ["1W"] * len(exs), ["0deg"] * len(exs))
            self.populate_source_controls(exs)
            self.log_message("Initial sources set to 1W/0deg.")
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
        try:
            names = self.hfss.get_excitations_name() or []
        except Exception:
            names = []
        if not names and self.created_ports:
            names = [f"{p}:1" for p in self.created_ports]

        def keyfn(s: str) -> int:
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
            sd = self._fetch_solution(
                expr,
                setup_candidates=["Setup1 : LastAdaptive", "Setup1:LastAdaptive", "Setup1 : Sweep1", "Setup1:Sweep1"],
                primary_sweep_variable=prim, variations=variations, context="Infinite Sphere1"
            )
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
        try:
            phi_vals = np.arange(-180.0, 180.0 + phi_step, phi_step)
            TH_list, G_list = None, []
            for phi in phi_vals:
                sd = self._fetch_solution(
                    "dB(GainTotal)",
                    setup_candidates=["Setup1 : LastAdaptive", "Setup1:LastAdaptive", "Setup1 : Sweep1", "Setup1:Sweep1"],
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
                        continue
                G_list.append(g)
            if TH_list is None or len(G_list) == 0:
                return None
            G = np.vstack(G_list).T  # (Ntheta, Nphi)
            PH = phi_vals[:G.shape[1]]
            TH = TH_list
            return TH, PH, G
        except Exception as e:
            self.log_message(f"3D grid error: {e}")
            return None

    # ------------- Simulação -------------
    def start_simulation_thread(self):
        """Thread para não travar a UI."""
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
            self.progress_bar.set(0.15)
            self.hfss.modeler.model_units = "mm"; self.log_message("Model units set to: mm")
            self._set_design_variables_common()

            # Geometria conforme esquema
            self.created_ports.clear()
            scheme = self.params["feed_scheme"]
            if scheme.startswith("Aperture"):
                geom = self._build_aperture_coupled_stack()
                self.progress_bar.set(0.35)
                # Região e radiação
                lambda0_mm = self.c / (self.params["sweep_start"] * 1e9) * 1000.0
                pad_mm = float(lambda0_mm) / 4.0
                region = self.hfss.modeler.create_region([pad_mm]*6, is_percentage=False)
                self.hfss.assign_radiation_boundary_to_objects(region)
                ground_for_rad = [geom["gnd_mid"].name]
            else:
                ground, patches = self._build_array_coax()
                self.progress_bar.set(0.35)
                lambda0_mm = self.c / (self.params["sweep_start"] * 1e9) * 1000.0
                pad_mm = float(lambda0_mm) / 4.0
                region = self.hfss.modeler.create_region([pad_mm]*6, is_percentage=False)
                self.hfss.assign_radiation_boundary_to_objects(region)
                ground_for_rad = [ground.name]

            # Esfera
            self._ensure_infinite_sphere("Infinite Sphere1")

            # Setup
            self.log_message("Creating simulation setup")
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02
            try:
                setup.props["SaveFields"] = False
                setup.props["SaveRadFields"] = True
            except Exception:
                pass

            # Sweep
            self.log_message(f"Creating frequency sweep: {self.params['sweep_type']}")
            stype = self.params["sweep_type"]
            try:
                try:
                    sw = setup.get_sweep("Sweep1");  sw.delete() if sw else None
                except Exception:
                    pass
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
            try:
                _ = self.hfss.validate_full_design()
            except Exception as e:
                self.log_message(f"Validation warning: {e}")

            self.log_message("Starting analysis")
            if self.save_project:
                self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])
            if self.stop_simulation:
                self.log_message("Simulation stopped by user"); return

            self._postprocess_after_solve()
            self.progress_bar.set(0.9)
            self.log_message("Processing results")
            self.analyze_and_mark_s11()
            self.refresh_patterns_only()
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

        for expr_tpl in [f"( {port_name},{port_name} )", f"({port_name},{port_name})"]:
            sd_db = self._fetch_solution(
                f"dB(S{expr_tpl})",
                setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1", "Setup1 : LastAdaptive"]
            )
            if sd_db and hasattr(sd_db, "primary_sweep_values"):
                f = np.asarray(sd_db.primary_sweep_values, dtype=float)
                y_db = self._shape_series(sd_db.data_real(), f.size)
                if y_db.size == f.size and f.size > 0:
                    sd_re = self._fetch_solution(f"re(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
                    sd_im = self._fetch_solution(f"im(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
                    reS = self._shape_series(sd_re.data_real(), f.size) if sd_re else None
                    imS = self._shape_series(sd_im.data_real(), f.size) if sd_im else None
                    return f, y_db, reS, imS

        expr_tpl = "(1,1)"
        sd_db = self._fetch_solution(
            f"dB(S{expr_tpl})",
            setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1", "Setup1 : LastAdaptive"]
        )
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
            self.ax_s11.clear(); self.ax_imp.clear()
            data = self._get_s11_curves()
            if not data:
                self.log_message("Solution Data failed to load. Check solution, context or expression.")
                self.canvas.draw(); return
            f, s11_db, reS, imS = data
            if f.size == 0 or s11_db.size == 0:
                self.log_message("S11 analysis aborted: empty curve.")
                self.canvas.draw(); return

            self.ax_s11.plot(f, s11_db, linewidth=2, label="S11 (dB)")
            self.ax_s11.axhline(y=-10, linestyle='--', alpha=0.7, label='-10 dB')
            self.ax_s11.set_xlabel("Frequency (GHz)")
            self.ax_s11.set_ylabel("S11 (dB)")
            self.ax_s11.set_title("S11 & VSWR")
            self.ax_s11.grid(True, alpha=0.5)

            s_abs = 10 ** (s11_db / 20.0)
            s_abs = np.clip(s_abs, 0, 0.999999)
            vswr = (1 + s_abs) / (1 - s_abs)
            ax_v = self.ax_s11.twinx()
            ax_v.plot(f, vswr, linestyle='--', alpha=0.8, label='VSWR')
            ax_v.set_ylabel("VSWR")

            idx_min = int(np.argmin(s11_db))
            f_res = float(f[idx_min]); s11_min_db = float(s11_db[idx_min])
            self.ax_s11.scatter([f_res], [s11_min_db], s=45, marker="o", zorder=5)
            self.ax_s11.annotate(f"f_res={f_res:.4g} GHz\nS11={s11_min_db:.2f} dB",
                                 (f_res, s11_min_db), textcoords="offset points", xytext=(8, -16))
            cf = float(self.params["frequency"])
            self.ax_s11.axvline(x=cf, linestyle=':', alpha=0.7, color='r', label=f"f0={cf:g} GHz")

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

            self.last_s11_analysis = {"f": f, "s11_db": s11_db, "vswr": vswr,
                                      "Zmag": Zmag, "f_res": f_res,
                                      "R": R, "X": X}

            if R is not None and X is not None:
                self.result_label.configure(text=f"Min @ {f_res:.4g} GHz, S11={s11_min_db:.2f} dB, Z≈{R:.1f} + j{X:.1f} Ω")
            else:
                self.result_label.configure(text=f"Min @ {f_res:.4g} GHz, S11={s11_min_db:.2f} dB")

            self.canvas.draw()
        except Exception as e:
            self.log_message(f"Analyze S11 error: {e}\nTraceback: {traceback.format_exc()}")

    # ------------- Padrões / 3D -------------
    def refresh_patterns_only(self):
        try:
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

            grid = self._get_gain_3d_grid(f0, theta_step=self.params["theta_step"], phi_step=self.params["phi_step"])
            if grid is not None:
                TH_deg, PH_deg, Gdb = grid
                self.grid3d = grid
                TH = np.deg2rad(TH_deg)[:, None] * np.ones((1, Gdb.shape[1]))
                PH = np.deg2rad(PH_deg)[None, :] * np.ones((Gdb.shape[0], 1))
                Glin = 10 ** (Gdb / 20.0)
                Glin = Glin - np.min(Glin)
                if np.max(Glin) > 0:
                    Glin = Glin / np.max(Glin)
                R = 0.2 + 0.8 * Glin
                X = R * np.sin(TH) * np.cos(PH)
                Y = R * np.sin(TH) * np.sin(PH)
                Z = R * np.cos(TH)
                self.ax_3d.plot_surface(X, Y, Z, rstride=1, cstride=1, linewidth=0, antialiased=True,
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
        for child in self.src_frame.winfo_children():
            child.destroy()
        self.source_controls.clear()
        if not excitations:
            ctk.CTkLabel(self.src_frame, text="No excitations found.").pack(padx=8, pady=6)
            return

        head = ctk.CTkFrame(self.src_frame); head.pack(fill="x", padx=8, pady=4)
        ctk.CTkLabel(head, text="Beamforming & Refresh", font=ctk.CTkFont(weight="bold")).pack(side="left")
        grid = ctk.CTkFrame(self.src_frame); grid.pack(fill="x", padx=8, pady=6)

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
            pvals, phvals = [], []
            for i, ex in enumerate(exs, start=1):
                ctrl = self.source_controls.get(ex)
                if not ctrl:
                    continue
                try:
                    pw = float(ctrl["power"].get())
                except Exception:
                    pw = 1.0
                ph = float(ctrl["phase"].get())
                pvals.append(f"{pw}W")
                phvals.append(f"{ph}deg")
            self._edit_sources_with_vars(exs, pvals, phvals)
            self.refresh_patterns_only()
        except Exception as e:
            self.log_message(f"Apply sources error: {e}\nTraceback: {traceback.format_exc()}")

    def toggle_auto_refresh(self):
        if self.auto_refresh_var.get():
            self.schedule_auto_refresh()
        else:
            if self.auto_refresh_job:
                try:
                    self.window.after_cancel(self.auto_refresh_job)
                except Exception:
                    pass
                self.auto_refresh_job = None

    def schedule_auto_refresh(self):
        self.refresh_patterns_only()
        if self.auto_refresh_var.get():
            self.auto_refresh_job = self.window.after(1500, self.schedule_auto_refresh)

    # ------------- Exportações -------------
    def export_csv(self):
        try:
            if self.last_s11_analysis is not None:
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

    # ------------- Cleanup / Persistência -------------
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
                    self.log_message(f"Error cleaning temp files: {e}")
        except Exception as e:
            self.log_message(f"Error during cleanup: {e}")

    def on_closing(self):
        self.log_message("Application closing...")
        self.cleanup()
        self.window.quit(); self.window.destroy()

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
                        widget.delete(0, "end"); widget.insert(0, str(self.params[key]))
                    elif isinstance(widget, ctk.StringVar):
                        widget.set(self.params[key])
                    elif isinstance(widget, ctk.BooleanVar):
                        widget.set(self.params[key])
            self.patches_label.configure(text=f"Number of Patches: {self.calculated_params['num_patches']}")
            self.rows_cols_label.configure(text=f"Configuration: {self.calculated_params['rows']} x {self.calculated_params['cols']}")
            self.spacing_label.configure(text=f"Spacing: {self.calculated_params['spacing']:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(
                text=f"Patch Dimensions: {self.calculated_params['patch_length']:.2f} x "
                     f"{self.calculated_params['patch_width']:.2f} mm"
            )
            self.lambda_label.configure(text=f"Guided Wavelength: {self.calculated_params['lambda_g']:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset (y): {self.calculated_params['feed_offset']:.2f} mm")
            self.substrate_dims_label.configure(
                text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x "
                     f"{self.calculated_params['substrate_length']:.2f} mm"
            )
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
