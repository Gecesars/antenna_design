# -*- coding: utf-8 -*-
"""
Modern Patch Antenna Designer (v4.3) - Versão Final
---------------------------------------------------------
Implementações e Correções Finais:
- Corrigido erro de "variable not found" ao unir objetos, otimizando a sequência de criação de geometria.
- Corrigida a extração de resultados (S-Parameters e Gain) para usar expressões robustas ('st' para terminais) e contextos corretos.
- Reorganizada a interface gráfica com abas dedicadas para "S-Parameters" e "Radiation", melhorando a usabilidade.
- Refinada a lógica de pós-processamento para garantir que os dados sejam sempre obtidos da solução correta.
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
from typing import Tuple, List, Optional, Dict, Any

import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.mplot3d import Axes3D
import customtkinter as ctk
from tkinter import messagebox, ttk

from ansys.aedt.core import Desktop, Hfss

# ---------- Aparência ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ModernPatchAntennaDesigner:
    """Aplicativo GUI para dimensionamento e simulação de patch array em HFSS."""

    # ---------------- Inicialização ----------------
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
        self.save_project = False
        self.simulation_running = False

        # Dados em memória
        self.last_s11_analysis = None
        self.theta_cut = None
        self.phi_cut = None
        self.grid3d = None
        self.auto_refresh_job = None
        
        # Otimização
        self.original_params = {}
        self.optimized = False
        self.optimization_history = []
        self.original_s11_data = None
        self.original_theta_data = None
        self.original_phi_data = None

        # Parâmetros do usuário (default)
        self.params = {
            "frequency": 10.0, "gain": 12.0, "sweep_start": 8.0, "sweep_stop": 12.0,
            "cores": 4, "aedt_version": "2024.2", "non_graphical": False, "spacing_type": "lambda/2",
            "substrate_material": "Duroid (tm)", "substrate_thickness": 0.5, "metal_thickness": 0.035,
            "er": 2.2, "tan_d": 0.0009, "feed_position": "inset", "feed_rel_x": 0.485,
            "probe_radius": 0.40, "coax_ba_ratio": 2.3, "coax_wall_thickness": 0.20,
            "coax_port_length": 3.0, "antipad_clearance": 0.10, "sweep_type": "Interpolating",
            "sweep_step": 0.02, "theta_step": 10.0, "phi_step": 10.0
        }

        # Parâmetros calculados
        self.calculated_params = {
            "num_patches": 4, "spacing": 0.0, "patch_length": 9.57, "patch_width": 9.25,
            "rows": 2, "cols": 2, "lambda_g": 0.0, "feed_offset": 2.0,
            "substrate_width": 0.0, "substrate_length": 0.0
        }

        self.c = 299792458.0
        self.setup_gui()
        self._style_ttk_treeview()

    # ---------------- Configuração da GUI ----------------
    def setup_gui(self):
        """Constroi a janela principal e abas com layout profissional."""
        self.window = ctk.CTk()
        self.window.title("Modern Patch Antenna Array Designer (v4.3)")
        self.window.geometry("1600x1000")
        
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.window, height=80, fg_color=("gray90", "gray15"))
        header.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 5))
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)
        
        logo_frame = ctk.CTkFrame(header, width=60, height=60, fg_color=("gray80", "gray25"))
        logo_frame.grid(row=0, column=0, padx=15, pady=10, sticky="w")
        logo_frame.grid_propagate(False)
        ctk.CTkLabel(logo_frame, text="ANT", font=ctk.CTkFont(size=20, weight="bold")).pack(expand=True)
        
        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkLabel(title_frame, text="Modern Patch Antenna Array Designer", font=ctk.CTkFont(size=24, weight="bold"), text_color=("gray10", "gray90")).pack(anchor="w")
        ctk.CTkLabel(title_frame, text="Professional RF Design Tool", font=ctk.CTkFont(size=14), text_color=("gray40", "gray60")).pack(anchor="w")

        status_frame = ctk.CTkFrame(header, fg_color="transparent")
        status_frame.grid(row=0, column=2, padx=15, pady=10, sticky="e")
        self.quick_status = ctk.CTkLabel(status_frame, text="Ready", font=ctk.CTkFont(weight="bold"), fg_color=("gray85", "gray25"), corner_radius=8)
        self.quick_status.pack(padx=5, pady=5, ipadx=10, ipady=5)

        self.tabview = ctk.CTkTabview(self.window, fg_color=("gray95", "gray10"))
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        tabs = ["Design", "Simulation", "S-Parameters", "Radiation", "Optimization", "Log"]
        for name in tabs:
            tab = self.tabview.add(name)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

        self.setup_design_tab()
        self.setup_simulation_tab()
        self.setup_sparameters_tab()
        self.setup_radiation_tab()
        self.setup_optimization_tab()
        self.setup_log_tab()

        status = ctk.CTkFrame(self.window, height=40, fg_color=("gray92", "gray18"))
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        status.grid_columnconfigure(0, weight=1)
        
        self.status_label = ctk.CTkLabel(status, text="Ready", font=ctk.CTkFont(weight="bold"), anchor="w", text_color=("gray30", "gray70"))
        self.status_label.grid(row=0, column=0, padx=15, pady=6, sticky="w")
        
        version_label = ctk.CTkLabel(status, text="v4.3 © 2025 RF Design Suite", font=ctk.CTkFont(size=12), text_color=("gray40", "gray60"))
        version_label.grid(row=0, column=1, padx=15, pady=6, sticky="e")

        self.process_log_queue()

    def create_section(self, parent, title, description=None, row=0, column=0, padx=10, pady=10, colspan=1):
        """Cria um frame com título e descrição para organizar a UI."""
        section = ctk.CTkFrame(parent, fg_color=("gray97", "gray12"), corner_radius=8)
        section.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady, columnspan=colspan)
        section.grid_columnconfigure(0, weight=1)
        
        header_frame = ctk.CTkFrame(section, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        header_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(header_frame, text=title, font=ctk.CTkFont(size=16, weight="bold"), text_color=("gray20", "gray80")).grid(row=0, column=0, sticky="w")
        
        if description:
            ctk.CTkLabel(header_frame, text=description, font=ctk.CTkFont(size=12), text_color=("gray40", "gray60")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        ctk.CTkFrame(section, height=2, fg_color=("gray80", "gray25")).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        
        return section

    def setup_design_tab(self):
        """Aba de parâmetros de projeto com layout profissional."""
        tab = self.tabview.tab("Design")
        main = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=1)

        self.entries = []
        def add_entry(parent, label, key, value, row, **kwargs):
            unit, tooltip, combo, check = kwargs.get("unit"), kwargs.get("tooltip"), kwargs.get("combo"), kwargs.get("check")
            field_frame = ctk.CTkFrame(parent, fg_color="transparent")
            field_frame.grid(row=row, column=0, sticky="ew", padx=15, pady=4)
            field_frame.grid_columnconfigure(1, weight=1)
            label_text = f"{label} ({unit}):" if unit else f"{label}:"
            lbl = ctk.CTkLabel(field_frame, text=label_text, font=ctk.CTkFont(size=12, weight="bold"), anchor="w", text_color=("gray30", "gray70"))
            lbl.grid(row=0, column=0, sticky="w", padx=(0, 10))
            widget_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
            widget_frame.grid(row=0, column=1, sticky="ew")
            widget_frame.grid_columnconfigure(0, weight=1)
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(widget_frame, values=combo, variable=var)
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(widget_frame, text="", variable=var)
                self.entries.append((key, var))
            else:
                var = ctk.StringVar(value=str(value))
                widget = ctk.CTkEntry(widget_frame, textvariable=var)
                self.entries.append((key, var))
            widget.grid(row=0, column=0, sticky="ew")
            if tooltip:
                info_btn = ctk.CTkLabel(widget_frame, text="ⓘ", text_color=("gray50", "gray50"), cursor="hand2")
                info_btn.grid(row=0, column=1, padx=(5, 0))
                info_btn.bind("<Enter>", lambda e, t=tooltip: self.show_tooltip(e, t))
                info_btn.bind("<Leave>", self.hide_tooltip)

        sec_ant = self.create_section(main, "Antenna Parameters", "Fundamental design parameters", 0, 0)
        row_idx = 2
        add_entry(sec_ant, "Central Frequency", "frequency", self.params["frequency"], row_idx, unit="GHz", tooltip="Operating frequency")
        add_entry(sec_ant, "Desired Gain", "gain", self.params["gain"], row_idx+1, unit="dBi", tooltip="Target gain")
        add_entry(sec_ant, "Sweep Start", "sweep_start", self.params["sweep_start"], row_idx+2, unit="GHz")
        add_entry(sec_ant, "Sweep Stop", "sweep_stop", self.params["sweep_stop"], row_idx+3, unit="GHz")
        add_entry(sec_ant, "Patch Spacing", "spacing_type", self.params["spacing_type"], row_idx+4, combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        sec_sub = self.create_section(main, "Substrate Parameters", "Material properties", 0, 1)
        row_idx = 2
        add_entry(sec_sub, "Substrate Material", "substrate_material", self.params["substrate_material"], row_idx, combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        add_entry(sec_sub, "Relative Permittivity (εr)", "er", self.params["er"], row_idx+1)
        add_entry(sec_sub, "Loss Tangent (tan δ)", "tan_d", self.params["tan_d"], row_idx+2)
        add_entry(sec_sub, "Substrate Thickness", "substrate_thickness", self.params["substrate_thickness"], row_idx+3, unit="mm")
        add_entry(sec_sub, "Metal Thickness", "metal_thickness", self.params["metal_thickness"], row_idx+4, unit="mm")

        sec_coax = self.create_section(main, "Feed Parameters", "Coaxial feed configuration", 1, 0)
        row_idx = 2
        add_entry(sec_coax, "Feed Position Type", "feed_position", self.params["feed_position"], row_idx, combo=["inset", "edge"])
        add_entry(sec_coax, "Feed Relative X Position", "feed_rel_x", self.params["feed_rel_x"], row_idx+1)
        add_entry(sec_coax, "Probe Radius (a)", "probe_radius", self.params["probe_radius"], row_idx+2, unit="mm")
        add_entry(sec_coax, "b/a Ratio", "coax_ba_ratio", self.params["coax_ba_ratio"], row_idx+3, tooltip="≈2.3 for 50Ω")
        add_entry(sec_coax, "Port Length", "coax_port_length", self.params["coax_port_length"], row_idx+4, unit="mm")
        add_entry(sec_coax, "Anti-Pad Clearance", "antipad_clearance", self.params["antipad_clearance"], row_idx+5, unit="mm")

        sec_sim = self.create_section(main, "Simulation Settings", "Solution configuration", 1, 1)
        row_idx = 2
        add_entry(sec_sim, "CPU Cores", "cores", self.params["cores"], row_idx)
        add_entry(sec_sim, "Show HFSS UI", "show_gui", not self.params["non_graphical"], row_idx+1, check=True)
        add_entry(sec_sim, "Save Project", "save_project", self.save_project, row_idx+2, check=True)
        add_entry(sec_sim, "Sweep Type", "sweep_type", self.params["sweep_type"], row_idx+3, combo=["Discrete", "Interpolating", "Fast"])
        add_entry(sec_sim, "Discrete Step", "sweep_step", self.params["sweep_step"], row_idx+4, unit="GHz")
        add_entry(sec_sim, "3D Theta Step", "theta_step", self.params["theta_step"], row_idx+5, unit="deg")
        add_entry(sec_sim, "3D Phi Step", "phi_step", self.params["phi_step"], row_idx+6, unit="deg")

        sec_calc = self.create_section(main, "Calculated Parameters", "Derived design values", 2, 0, colspan=2)
        calc_grid = ctk.CTkFrame(sec_calc, fg_color="transparent")
        calc_grid.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        for i in range(2): calc_grid.grid_columnconfigure(i, weight=1)
        
        self.patches_label = ctk.CTkLabel(calc_grid, text="Number of Patches: --", anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        self.patches_label.grid(row=0, column=0, sticky="w", pady=6)
        self.rows_cols_label = ctk.CTkLabel(calc_grid, text="Configuration: -- x --", anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=6)
        self.spacing_label = ctk.CTkLabel(calc_grid, text="Spacing: -- mm", anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        self.spacing_label.grid(row=1, column=0, sticky="w", pady=6)
        self.dimensions_label = ctk.CTkLabel(calc_grid, text="Patch Dimensions: -- x -- mm", anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        self.dimensions_label.grid(row=1, column=1, sticky="w", pady=6)
        self.lambda_label = ctk.CTkLabel(calc_grid, text="Guided Wavelength: -- mm", anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        self.lambda_label.grid(row=2, column=0, sticky="w", pady=6)
        self.feed_offset_label = ctk.CTkLabel(calc_grid, text="Feed Offset (y): -- mm", anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        self.feed_offset_label.grid(row=2, column=1, sticky="w", pady=6)
        self.substrate_dims_label = ctk.CTkLabel(calc_grid, text="Substrate Dimensions: -- x -- mm", anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        self.substrate_dims_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=6)

        btn_frame = ctk.CTkFrame(sec_calc, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=12)
        for i in range(3): btn_frame.grid_columnconfigure(i, weight=1)
        
        ctk.CTkButton(btn_frame, text="Calculate Parameters", command=self.calculate_parameters, fg_color="#2E8B57", hover_color="#3CB371").grid(row=0, column=0, padx=8, sticky="ew")
        ctk.CTkButton(btn_frame, text="Save Parameters", command=self.save_parameters, fg_color="#4169E1", hover_color="#6495ED").grid(row=0, column=1, padx=8, sticky="ew")
        ctk.CTkButton(btn_frame, text="Load Parameters", command=self.load_parameters, fg_color="#FF8C00", hover_color="#FFA500").grid(row=0, column=2, padx=8, sticky="ew")

    def setup_simulation_tab(self):
        """Aba de simulação com layout profissional."""
        tab = self.tabview.tab("Simulation")
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        
        sec_control = self.create_section(main, "Simulation Control", "Run and monitor simulations", 0, 0)
        
        btn_frame = ctk.CTkFrame(sec_control, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=10)
        btn_frame.grid_columnconfigure(0, weight=1)
        
        self.run_button = ctk.CTkButton(btn_frame, text="▶ Run Simulation", command=self.start_simulation_thread, fg_color="#2E8B57", hover_color="#3CB371", height=40, font=ctk.CTkFont(size=14, weight="bold"))
        self.run_button.grid(row=0, column=0, padx=8, sticky="ew")
        
        status_frame = ctk.CTkFrame(sec_control, fg_color=("gray92", "gray18"), corner_radius=8)
        status_frame.grid(row=4, column=0, sticky="ew", padx=15, pady=10)
        status_frame.grid_columnconfigure(0, weight=1)
        
        self.sim_status_label = ctk.CTkLabel(status_frame, text="Simulation not started", font=ctk.CTkFont(weight="bold"), text_color=("gray30", "gray70"))
        self.sim_status_label.grid(row=0, column=0, padx=15, pady=12)

    def setup_sparameters_tab(self):
        """Aba de resultados para S-Parameters e Impedância."""
        tab = self.tabview.tab("S-Parameters")
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)

        graph_frame = ctk.CTkFrame(main, fg_color=("gray96", "gray14"), corner_radius=10)
        graph_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)

        self.fig_s, (self.ax_s11, self.ax_imp) = plt.subplots(1, 2, figsize=(12, 6))
        
        self.canvas_s = FigureCanvasTkAgg(self.fig_s, master=graph_frame)
        self.canvas_s.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        
        control_frame = ctk.CTkFrame(main)
        control_frame.grid(row=1, column=0, sticky="ew", pady=(10,0))
        
        self.result_label = ctk.CTkLabel(control_frame, text="Run simulation to see results", font=ctk.CTkFont(weight="bold"))
        self.result_label.pack(pady=5)
        
        btn_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        btn_frame.pack(pady=5)
        ctk.CTkButton(btn_frame, text="Analyze S11", command=self.analyze_and_mark_s11).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Export PNG", command=lambda: self.export_png(self.fig_s, "sparameters.png")).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Export CSV", command=self.export_csv).pack(side="left", padx=5)
        
    def setup_radiation_tab(self):
        """Aba de resultados para Padrões de Irradiação e Beamforming."""
        tab = self.tabview.tab("Radiation")
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)

        graph_frame = ctk.CTkFrame(main, fg_color=("gray96", "gray14"), corner_radius=10)
        graph_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)

        self.fig_rad = plt.figure(figsize=(14, 8))
        gs = self.fig_rad.add_gridspec(2, 2, hspace=0.35, wspace=0.25)
        self.ax_th = self.fig_rad.add_subplot(gs[0, 0])
        self.ax_ph = self.fig_rad.add_subplot(gs[0, 1])
        self.ax_3d = self.fig_rad.add_subplot(gs[1, :], projection='3d')
        
        self.canvas_rad = FigureCanvasTkAgg(self.fig_rad, master=graph_frame)
        self.canvas_rad.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        
        control_frame = ctk.CTkFrame(main)
        control_frame.grid(row=1, column=0, sticky="ew", pady=(10,0))

        self.src_frame = self.create_section(control_frame, "Beamforming Control", row=0, column=0, padx=0)
        
        btn_frame = ctk.CTkFrame(self.src_frame, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=5, padx=10)
        ctk.CTkButton(btn_frame, text="Apply Sources & Refresh", command=self.apply_sources_from_ui).pack(side="left", padx=5)
        self.auto_refresh_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(btn_frame, text="Auto-refresh", variable=self.auto_refresh_var, command=self.toggle_auto_refresh).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Export PNG", command=lambda: self.export_png(self.fig_rad, "radiation_patterns.png")).pack(side="left", padx=5)

    def setup_optimization_tab(self):
        """Aba de otimização com layout profissional."""
        tab = self.tabview.tab("Optimization")
        tab.grid_rowconfigure(1, weight=1)

        sec_opt = self.create_section(tab, "Optimization Control", "Tune design for better performance", 0, 0)
        self.opt_status_label = ctk.CTkLabel(sec_opt, text="No optimization performed yet", font=ctk.CTkFont(weight="bold"))
        self.opt_status_label.grid(row=2, column=0, padx=15, pady=12, sticky="ew")

        btn_frame = ctk.CTkFrame(sec_opt, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=15)
        for i in range(3): btn_frame.grid_columnconfigure(i, weight=1)
        ctk.CTkButton(btn_frame, text="Analyze & Optimize", command=self.analyze_and_optimize, fg_color="#2E8B57", hover_color="#3CB371").grid(row=0, column=0, padx=8, sticky="ew")
        ctk.CTkButton(btn_frame, text="Reset to Original", command=self.reset_to_original, fg_color="#DC143C", hover_color="#FF4500").grid(row=0, column=1, padx=8, sticky="ew")
        ctk.CTkButton(btn_frame, text="View History", command=self.view_optimization_history).grid(row=0, column=2, padx=8, sticky="ew")
        
        sec_history = self.create_section(tab, "Optimization History", "Track design changes", 1, 0)
        sec_history.grid_rowconfigure(2, weight=1)
        self.history_frame = ctk.CTkFrame(sec_history, fg_color=("gray92", "gray18"))
        self.history_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=15)
        self.history_frame.grid_columnconfigure(0, weight=1)
        self.history_frame.grid_rowconfigure(0, weight=1)
        ctk.CTkLabel(self.history_frame, text="Optimization history will be displayed here.").pack(expand=True)

    def setup_log_tab(self):
        """Aba de log com interface profissional."""
        tab = self.tabview.tab("Log")
        log_frame = ctk.CTkFrame(tab, fg_color="transparent")
        log_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        self.log_text = ctk.CTkTextbox(log_frame, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_text.pack(fill="both", expand=True)
        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(5,10), padx=5)
        for i in range(2): btn_frame.grid_columnconfigure(i, weight=1)
        ctk.CTkButton(btn_frame, text="Clear Log", command=self.clear_log, fg_color="#DC143C", hover_color="#FF4500").grid(row=0, column=0, padx=8, sticky="ew")
        ctk.CTkButton(btn_frame, text="Save Log", command=self.save_log).grid(row=0, column=1, padx=8, sticky="ew")

    # ------------- Utilitários de GUI e Estilo -------------
    def _style_plots(self, fig, axes_3d=[]):
        """Aplica estilo dinâmico aos gráficos Matplotlib baseado no tema CTk."""
        is_dark = ctk.get_appearance_mode() == "Dark"
        face_color = '#2B2B2B' if is_dark else '#F9F9FA'
        text_color = 'white' if is_dark else 'black'
        grid_color = '#404040' if is_dark else '#D3D3D3'

        fig.patch.set_facecolor(face_color)
        for ax in fig.get_axes():
            if ax in axes_3d:
                ax.set_facecolor(face_color)
                ax.tick_params(colors=text_color, which='both', pad=-3)
                for axis in [ax.xaxis, ax.yaxis, ax.zaxis]:
                    axis.label.set_color(text_color)
                    axis.set_pane_color((0,0,0,0))
                ax.title.set_color(text_color)
            else:
                ax.set_facecolor(face_color)
                ax.tick_params(colors=text_color, which='both')
                ax.xaxis.label.set_color(text_color)
                ax.yaxis.label.set_color(text_color)
                ax.title.set_color(text_color)
                for spine in ax.spines.values():
                    spine.set_edgecolor(text_color)
                ax.grid(True, color=grid_color, linestyle='--', linewidth=0.5)

    def _style_ttk_treeview(self):
        """Aplica um estilo moderno ao ttk.TreeView para combinar com o tema do CTk."""
        style = ttk.Style()
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg_color, fg_color = ("#2B2B2B", "#DCE4EE") if is_dark else ("#F9F9FA", "#333333")
        header_bg = "#303030" if is_dark else "#EAEAEA"
        selected_bg = "#1F6AA5"
        style.theme_use("default")
        style.configure("Treeview", background=bg_color, foreground=fg_color, fieldbackground=bg_color, borderwidth=0, rowheight=25)
        style.map('Treeview', background=[('selected', selected_bg)])
        style.configure("Treeview.Heading", background=header_bg, foreground=fg_color, relief="flat", font=('Calibri', 10, 'bold'))
        style.map("Treeview.Heading", background=[('active', '#3C3C3C' if is_dark else '#DCDCDC')])

    def show_tooltip(self, event, text):
        if hasattr(self, 'current_tooltip') and self.current_tooltip: self.current_tooltip.destroy()
        tooltip = ctk.CTkToplevel(self.window)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{event.x_root+15}+{event.y_root+10}")
        label = ctk.CTkLabel(tooltip, text=text, fg_color=("#EAEAEA", "#333333"), text_color=("#333333", "#EAEAEA"),
                             corner_radius=5, justify="left", wraplength=300, font=("Calibri", 12))
        label.pack(padx=8, pady=5)
        self.current_tooltip = tooltip
        tooltip.after(5000, tooltip.destroy)

    def hide_tooltip(self, event):
        if hasattr(self, 'current_tooltip') and self.current_tooltip:
            self.current_tooltip.destroy()
            self.current_tooltip = None


    def save_project_toggle(self):
        """Alterna o estado de salvamento do projeto."""
        self.save_project = not self.save_project
        status = "ON" if self.save_project else "OFF"
        self.log_message(f"Project saving {status}")
        self.status_label.configure(text=f"Project saving {status}")

    def export_report(self):
        """(Placeholder) Exporta um relatório completo em PDF."""
        self.log_message("Funcionalidade de exportar relatório ainda não implementada.")
        messagebox.showinfo("Info", "A funcionalidade de exportar relatório será implementada em uma versão futura.")

    def compare_results(self):
        """Compara resultados entre a simulação original e a otimizada."""
        if not self.optimized or not self.original_s11_data:
            messagebox.showinfo("Compare Results", "Nenhuma otimização foi realizada para comparar.")
            return
        self.log_message("Exibindo comparação entre resultados originais e otimizados.")
        self.update_s_plots(compare_mode=True)
        self.update_radiation_plots(compare_mode=True)
        self.analyze_and_mark_s11()

    def view_optimization_history(self):
        """Exibe o histórico de otimização em uma janela separada usando ttk.TreeView."""
        if not self.optimization_history:
            messagebox.showinfo("Optimization History", "Nenhum histórico de otimização disponível.")
            return
            
        history_window = ctk.CTkToplevel(self.window)
        history_window.title("Optimization History")
        history_window.geometry("850x400")
        history_window.grab_set()
        
        ctk.CTkLabel(history_window, text="Optimization History", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        # Limpa o frame antes de adicionar a nova tabela
        for widget in self.history_frame.winfo_children():
            widget.destroy()

        columns = ("iteration", "resonant_freq", "target_freq", "error_percent", "min_s11", "scaling_factor")
        tree = ttk.Treeview(self.history_frame, columns=columns, show="headings")
        
        vsb = ttk.Scrollbar(self.history_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(self.history_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        for col in columns:
            tree.heading(col, text=col.replace("_", " ").title())
            tree.column(col, width=130, anchor='center')
        
        for record in self.optimization_history:
            tree.insert("", "end", values=(
                record["iteration"], f"{record['resonant_freq']:.3f}", f"{record['target_freq']:.3f}",
                f"{record['error_percent']:.1f}", f"{record['min_s11']:.2f}", f"{record['scaling_factor']:.4f}"
            ))
        
        ctk.CTkButton(history_window, text="Close", command=history_window.destroy).pack(pady=10)

    def update_quick_status(self, status, color=None):
        """Atualiza o status rápido no header."""
        colors = {"ready": ("gray85", "gray25"), "running": ("#FFA500", "#CC8400"),
                  "success": ("#2E8B57", "#3CB371"), "error": ("#DC143C", "#FF4500")}
        if color in colors:
            self.quick_status.configure(fg_color=colors[color])
        self.quick_status.configure(text=status)

    # ------------- Utilitários de Log -------------
    def log_message(self, message: str):
        """Enfileira uma mensagem para o textbox de log com carimbo de hora."""
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")

    def process_log_queue(self):
        """Consumidor assíncrono da fila de log."""
        try:
            while not self.log_queue.empty():
                msg = self.log_queue.get_nowait()
                self.log_text.insert("end", msg)
                self.log_text.see("end")
        finally:
            if self.window.winfo_exists():
                self.window.after(100, self.process_log_queue)

    def clear_log(self):
        self.log_text.delete("1.0", "end")
        self.log_message("Log cleared.")

    def save_log(self):
        from tkinter.filedialog import asksaveasfilename
        filepath = asksaveasfilename(defaultextension=".log", filetypes=[("Log files", "*.log"), ("All files", "*.*")],
                                     initialfile=f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        if not filepath: return
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(self.log_text.get("1.0", "end"))
            self.log_message(f"Log saved to {filepath}")
        except Exception as e:
            self.log_message(f"Error saving log: {e}")

    # ------------- Lógica Principal e Execução da Simulação -------------
    def start_simulation_thread(self):
        """Inicia a simulação em uma thread separada para não bloquear a GUI."""
        if self.simulation_running:
            messagebox.showwarning("Warning", "A simulation is already in progress.")
            return

        self.simulation_running = True
        self.run_button.configure(state="disabled", text="■ Running...")
        self.sim_status_label.configure(text="Simulation starting...")
        self.update_quick_status("Running...", "running")
        self.log_message("Starting simulation thread...")

        thread = threading.Thread(target=self._run_simulation_task)
        thread.daemon = True
        thread.start()

    def _run_simulation_task(self):
        """Task de simulação executada em uma thread separada, com todas as chamadas corrigidas."""
        try:
            self.log_message("Validating and calculating parameters...")
            self.window.after(0, lambda: self.sim_status_label.configure(text="Calculating parameters..."))
            self.calculate_parameters()

            self.window.after(0, lambda: self.sim_status_label.configure(text="Initializing AEDT..."))
            self._open_or_create_project()

            self.log_message("Creating geometry and boundaries...")
            self.window.after(0, lambda: self.sim_status_label.configure(text="Creating geometry..."))
            self._create_geometry_and_boundaries()

            self.log_message("Creating analysis setup...")
            self.window.after(0, lambda: self.sim_status_label.configure(text="Creating analysis setup..."))
            self._create_analysis_setup()

            self.log_message("Starting analysis...")
            self.window.after(0, lambda: self.sim_status_label.configure(text="Solving... (This may take a while)"))
            self.hfss.analyze(setup_name="Setup1")

            self.log_message("Performing post-solve setup for beamforming...")
            self.window.after(0, lambda: self.sim_status_label.configure(text="Post-processing..."))
            self._postprocess_after_solve()

            self.window.after(0, self._on_simulation_complete, True, None)

        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            self.log_message(error_msg)
            self.log_message(f"Traceback: {traceback.format_exc()}")
            self.window.after(0, self._on_simulation_complete, False, str(e))

    def _on_simulation_complete(self, success: bool, error_msg: Optional[str] = None):
        """Callback executado na thread principal da GUI após a simulação."""
        if success:
            self.log_message("Simulation completed successfully.")
            self.sim_status_label.configure(text="Simulation finished successfully. Fetching results...")
            self.update_quick_status("Success", "success")
            self.fetch_and_plot_results()
        else:
            self.log_message(f"Simulation failed: {error_msg}")
            self.sim_status_label.configure(text=f"Simulation failed: {error_msg}")
            self.update_quick_status("Error", "error")
            messagebox.showerror("Simulation Error", f"The simulation failed.\n\nDetails: {error_msg}")

        self.simulation_running = False
        self.run_button.configure(state="normal", text="▶ Run Simulation")
    
    # ------------- Métodos de Persistência e Exportação -------------
    def save_parameters(self):
        """Salva os parâmetros atuais da UI em um arquivo JSON."""
        if not self.get_parameters():
            self.log_message("Cannot save invalid parameters.")
            return
        from tkinter.filedialog import asksaveasfilename
        filepath = asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")], initialfile=f"config_{datetime.now().strftime('%Y%m%d')}.json")
        if not filepath: return
        try:
            with open(filepath, 'w') as f: json.dump(self.params, f, indent=4)
            self.log_message(f"Parameters saved to {filepath}")
            self.status_label.configure(text=f"Parameters saved to {os.path.basename(filepath)}")
        except Exception as e:
            self.log_message(f"Error saving parameters: {e}")

    def load_parameters(self):
        """Carrega parâmetros de um arquivo JSON e atualiza a UI."""
        from tkinter.filedialog import askopenfilename
        filepath = askopenfilename(filetypes=[("JSON files", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'r') as f: loaded_params = json.load(f)
            self.params.update(loaded_params)
            for key, var in self.entries:
                if key in self.params:
                    if key == "show_gui": var.set(not self.params.get("non_graphical", False))
                    elif key == "save_project": var.set(self.save_project)
                    else: var.set(self.params[key])
            self.log_message(f"Parameters loaded from {filepath}")
            self.status_label.configure(text=f"Loaded from {os.path.basename(filepath)}. Please re-calculate.")
            self.calculate_parameters()
        except Exception as e:
            self.log_message(f"Error loading parameters: {e}")
            
    def export_png(self, fig, filename):
        """Exporta a figura especificada como um arquivo PNG."""
        from tkinter.filedialog import asksaveasfilename
        filepath = asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")], initialfile=filename)
        if filepath:
            fig.savefig(filepath, dpi=300, facecolor=fig.get_facecolor())
            self.log_message(f"Figure saved to {filepath}")

    def export_csv(self):
        """Exporta os dados S11 para um arquivo CSV."""
        if not self.last_s11_analysis:
            messagebox.showinfo("Info", "No S11 data to export.")
            return
        from tkinter.filedialog import asksaveasfilename
        filepath = asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")], initialfile="s_parameters.csv")
        if filepath:
            data = self.last_s11_analysis
            header = "Frequency (GHz),S11 (dB),Real(Z),Imag(Z)"
            np.savetxt(filepath, np.vstack([data['f'], data['s11_db'], data['z_real'], data['z_imag']]).T, delimiter=',', header=header, comments='')
            self.log_message(f"S11 data saved to {filepath}")

    # ------------- Loop Principal e Encerramento -------------
    def run(self):
        """Inicia o loop principal da GUI e garante o cleanup ao sair."""
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.mainloop()

    def on_closing(self):
        """Lida com o fechamento da janela, liberando recursos do AEDT."""
        if self.simulation_running:
            if not messagebox.askyesno("Confirm Quit", "A simulation is still running. Are you sure you want to quit? This may corrupt the simulation."):
                return
        
        self.log_message("Closing application and releasing AEDT resources...")
        if self.hfss:
            try: 
                self.hfss.release_desktop(close_projects=True, close_desktop=True)
                self.log_message("AEDT Desktop released.")
            except Exception as e:
                self.log_message(f"Error releasing AEDT: {e}")
        self.window.destroy()

if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()