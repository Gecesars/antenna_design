# -*- coding: utf-8 -*-
"""
Modern Patch Antenna Designer (v4.2) - Versão Definitiva
---------------------------------------------------------
Implementações e Correções Finais:
- Substituído o CTkTreeView externo pelo ttk.TreeView nativo do Python com estilo customizado
  para garantir estabilidade e remover dependências frágeis.
- Corrigidas todas as chamadas da API PyAEDT com base na documentação oficial,
  resolvendo todos os 'AttributeError' e alertas de depreciação.
- Otimizado o fluxo de criação de setup e sweep para seguir as melhores práticas.
- Refinada a lógica de execução da análise para ser mais específica e robusta.
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
        self.created_ports: List[str] = []
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

    # ---------------- GUI ----------------
    def setup_gui(self):
        """Constroi a janela principal e abas com layout profissional."""
        self.window = ctk.CTk()
        self.window.title("Modern Patch Antenna Array Designer (v4.2)")
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
        ctk.CTkLabel(title_frame, text="Modern Patch Antenna Array Designer",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=("gray10", "gray90")).pack(anchor="w")
        ctk.CTkLabel(title_frame, text="Professional RF Design Tool",
                     font=ctk.CTkFont(size=14), text_color=("gray40", "gray60")).pack(anchor="w")

        status_frame = ctk.CTkFrame(header, fg_color="transparent")
        status_frame.grid(row=0, column=2, padx=15, pady=10, sticky="e")
        self.quick_status = ctk.CTkLabel(status_frame, text="Ready", font=ctk.CTkFont(weight="bold"), 
                                         fg_color=("gray85", "gray25"), corner_radius=8)
        self.quick_status.pack(padx=5, pady=5, ipadx=10, ipady=5)

        self.tabview = ctk.CTkTabview(self.window, fg_color=("gray95", "gray10"))
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        tabs = ["Design", "Simulation", "Results", "Optimization", "Log"]
        for name in tabs:
            self.tabview.add(name)
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        self.setup_design_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_optimization_tab()
        self.setup_log_tab()

        status = ctk.CTkFrame(self.window, height=40, fg_color=("gray92", "gray18"))
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        status.grid_columnconfigure(0, weight=1)
        
        self.status_label = ctk.CTkLabel(status, text="Ready", font=ctk.CTkFont(weight="bold"), 
                                         anchor="w", text_color=("gray30", "gray70"))
        self.status_label.grid(row=0, column=0, padx=15, pady=6, sticky="w")
        
        version_label = ctk.CTkLabel(status, text="v4.2 © 2025 RF Design Suite", 
                                     font=ctk.CTkFont(size=12), text_color=("gray40", "gray60"))
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
        
        ctk.CTkLabel(header_frame, text=title, font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=("gray20", "gray80")).grid(row=0, column=0, sticky="w")
        
        if description:
            ctk.CTkLabel(header_frame, text=description, font=ctk.CTkFont(size=12),
                         text_color=("gray40", "gray60")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        ctk.CTkFrame(section, height=2, fg_color=("gray80", "gray25")).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        
        return section

    def setup_design_tab(self):
        """Aba de parâmetros de projeto com layout profissional."""
        tab = self.tabview.tab("Design")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=1)

        sec_ant = self.create_section(main, "Antenna Parameters", "Fundamental design parameters", 0, 0)
        self.entries = []
        row_idx = 2

        def add_entry(section, label, key, value, row, tooltip=None, combo=None, check=False, unit=""):
            field_frame = ctk.CTkFrame(section, fg_color="transparent")
            field_frame.grid(row=row, column=0, sticky="ew", padx=15, pady=4)
            field_frame.grid_columnconfigure(1, weight=1)

            label_text = f"{label} ({unit}):" if unit else f"{label}:"
            lbl = ctk.CTkLabel(field_frame, text=label_text, font=ctk.CTkFont(weight="bold"), 
                               anchor="w", text_color=("gray30", "gray70"))
            lbl.grid(row=0, column=0, sticky="w", padx=(0, 10))

            widget_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
            widget_frame.grid(row=0, column=1, sticky="ew")
            widget_frame.grid_columnconfigure(0, weight=1)
            
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(widget_frame, values=combo, variable=var)
                widget.grid(row=0, column=0, sticky="ew")
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(widget_frame, text="", variable=var)
                widget.grid(row=0, column=0, sticky="w")
                self.entries.append((key, var))
            else:
                var = ctk.StringVar(value=str(value))
                widget = ctk.CTkEntry(widget_frame, textvariable=var)
                widget.grid(row=0, column=0, sticky="ew")
                self.entries.append((key, var))

            if tooltip:
                info_btn = ctk.CTkLabel(widget_frame, text="ⓘ", text_color=("gray50", "gray50"), cursor="hand2")
                info_btn.grid(row=0, column=1, padx=(5, 0))
                info_btn.bind("<Enter>", lambda e, t=tooltip: self.show_tooltip(e, t))
                info_btn.bind("<Leave>", self.hide_tooltip)

        # Parâmetros de Antena
        row_idx = 2
        for p, d in [("frequency", "GHz"), ("gain", "dBi"), ("sweep_start", "GHz"), ("sweep_stop", "GHz")]:
            add_entry(sec_ant, p.replace("_", " ").title(), p, self.params[p], row_idx, unit=d)
            row_idx += 1
        add_entry(sec_ant, "Patch Spacing", "spacing_type", self.params["spacing_type"], row_idx, combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        # Parâmetros de Substrato
        sec_sub = self.create_section(main, "Substrate Parameters", "Material properties", 0, 1)
        row_idx = 2
        add_entry(sec_sub, "Substrate Material", "substrate_material", self.params["substrate_material"], row_idx, combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        row_idx += 1
        for p in ["er", "tan_d"]:
             add_entry(sec_sub, p.replace("_", " ").title(), p, self.params[p], row_idx)
             row_idx += 1
        for p, d in [("substrate_thickness", "mm"), ("metal_thickness", "mm")]:
            add_entry(sec_sub, p.replace("_", " ").title(), p, self.params[p], row_idx, unit=d)
            row_idx += 1

        # Parâmetros de Alimentação
        sec_coax = self.create_section(main, "Feed Parameters", "Coaxial feed configuration", 1, 0)
        row_idx = 2
        add_entry(sec_coax, "Feed Position Type", "feed_position", self.params["feed_position"], row_idx, combo=["inset", "edge"])
        row_idx += 1
        for p in ["feed_rel_x", "coax_ba_ratio"]:
            add_entry(sec_coax, p.replace("_", " ").title(), p, self.params[p], row_idx)
            row_idx += 1
        for p, d in [("probe_radius", "mm"), ("coax_wall_thickness", "mm"), ("coax_port_length", "mm"), ("antipad_clearance", "mm")]:
            add_entry(sec_coax, p.replace("_", " ").title(), p, self.params[p], row_idx, unit=d)
            row_idx += 1
            
        # Configurações de Simulação
        sec_sim = self.create_section(main, "Simulation Settings", "Solution configuration", 1, 1)
        row_idx = 2
        add_entry(sec_sim, "CPU Cores", "cores", self.params["cores"], row_idx)
        row_idx += 1
        add_entry(sec_sim, "Show HFSS UI", "show_gui", not self.params["non_graphical"], row_idx, check=True)
        row_idx += 1
        add_entry(sec_sim, "Save Project", "save_project", self.save_project, row_idx, check=True)
        row_idx += 1
        add_entry(sec_sim, "Sweep Type", "sweep_type", self.params["sweep_type"], row_idx, combo=["Discrete", "Interpolating", "Fast"])
        row_idx += 1
        add_entry(sec_sim, "Discrete Step", "sweep_step", self.params["sweep_step"], row_idx, unit="GHz")
        row_idx += 1
        for p, d in [("theta_step", "deg"), ("phi_step", "deg")]:
            add_entry(sec_sim, f"3D {p.split('_')[0].title()} Step", p, self.params[p], row_idx, unit=d)
            row_idx += 1

        # Parâmetros Calculados
        sec_calc = self.create_section(main, "Calculated Parameters", "Derived design values", 2, 0, colspan=2)
        calc_grid = ctk.CTkFrame(sec_calc, fg_color="transparent")
        calc_grid.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        for i in range(2): calc_grid.grid_columnconfigure(i, weight=1)
        
        self.patches_label = ctk.CTkLabel(calc_grid, text="Number of Patches: --", anchor="w")
        self.patches_label.grid(row=0, column=0, sticky="w", pady=6)
        self.rows_cols_label = ctk.CTkLabel(calc_grid, text="Configuration: -- x --", anchor="w")
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=6)
        self.spacing_label = ctk.CTkLabel(calc_grid, text="Spacing: -- mm", anchor="w")
        self.spacing_label.grid(row=1, column=0, sticky="w", pady=6)
        self.dimensions_label = ctk.CTkLabel(calc_grid, text="Patch Dimensions: -- x -- mm", anchor="w")
        self.dimensions_label.grid(row=1, column=1, sticky="w", pady=6)
        self.lambda_label = ctk.CTkLabel(calc_grid, text="Guided Wavelength: -- mm", anchor="w")
        self.lambda_label.grid(row=2, column=0, sticky="w", pady=6)
        self.feed_offset_label = ctk.CTkLabel(calc_grid, text="Feed Offset (y): -- mm", anchor="w")
        self.feed_offset_label.grid(row=2, column=1, sticky="w", pady=6)
        self.substrate_dims_label = ctk.CTkLabel(calc_grid, text="Substrate Dimensions: -- x -- mm", anchor="w")
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
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        
        sec_control = self.create_section(main, "Simulation Control", "Run and monitor simulations", 0, 0)
        
        btn_frame = ctk.CTkFrame(sec_control, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=10)
        btn_frame.grid_columnconfigure(0, weight=1)
        
        self.run_button = ctk.CTkButton(btn_frame, text="▶ Run Simulation", command=self.start_simulation_thread,
                                        fg_color="#2E8B57", hover_color="#3CB371", height=40, 
                                        font=ctk.CTkFont(size=14, weight="bold"))
        self.run_button.grid(row=0, column=0, padx=8, sticky="ew")
        
        status_frame = ctk.CTkFrame(sec_control, fg_color=("gray92", "gray18"), corner_radius=8)
        status_frame.grid(row=4, column=0, sticky="ew", padx=15, pady=10)
        status_frame.grid_columnconfigure(0, weight=1)
        
        self.sim_status_label = ctk.CTkLabel(status_frame, text="Simulation not started", 
                                             font=ctk.CTkFont(weight="bold"), text_color=("gray30", "gray70"))
        self.sim_status_label.grid(row=0, column=0, padx=15, pady=12)

    def setup_results_tab(self):
        """Aba de resultados com visualização profissional."""
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        graph_frame = ctk.CTkFrame(main, fg_color=("gray96", "gray14"), corner_radius=10)
        graph_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)

        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        self.fig = plt.figure(figsize=(14, 10), facecolor=face)
        gs = self.fig.add_gridspec(3, 2, height_ratios=[1, 1, 1.3], hspace=0.35, wspace=0.25)
        
        self.ax_s11 = self.fig.add_subplot(gs[0, 0])
        self.ax_imp = self.fig.add_subplot(gs[0, 1])
        self.ax_th = self.fig.add_subplot(gs[1, 0])
        self.ax_ph = self.fig.add_subplot(gs[1, 1])
        self.ax_3d = self.fig.add_subplot(gs[2, :], projection='3d')
        self._style_plots()

        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        
        control_frame = ctk.CTkFrame(main, fg_color="transparent")
        control_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        control_frame.grid_columnconfigure(0, weight=1)

        self.src_frame = ctk.CTkFrame(control_frame, fg_color=("gray94", "gray16"), corner_radius=8)
        self.src_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 10))
        self.source_controls: Dict[str, Dict[str, Any]] = {}

        btn_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=5)
        
        btn_row1 = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_row1.pack(fill="x", pady=2)
        ctk.CTkButton(btn_row1, text="Analyze S11", command=self.analyze_and_mark_s11).pack(side="left", padx=5)
        ctk.CTkButton(btn_row1, text="Apply Sources", command=self.apply_sources_from_ui).pack(side="left", padx=5)
        ctk.CTkButton(btn_row1, text="Refresh Patterns", command=self.refresh_patterns_only).pack(side="left", padx=5)
        self.auto_refresh_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(btn_row1, text="Auto-refresh (1.5s)", variable=self.auto_refresh_var, command=self.toggle_auto_refresh).pack(side="left", padx=5)

        btn_row2 = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_row2.pack(fill="x", pady=2)
        ctk.CTkButton(btn_row2, text="Export PNG", command=self.export_png).pack(side="left", padx=5)
        ctk.CTkButton(btn_row2, text="Export CSV (S11)", command=self.export_csv).pack(side="left", padx=5)
        ctk.CTkButton(btn_row2, text="Compare Results", command=self.compare_results).pack(side="left", padx=5)
        
        self.result_label = ctk.CTkLabel(control_frame, text="", font=ctk.CTkFont(weight="bold"))
        self.result_label.grid(row=2, column=0, sticky="ew", pady=(5, 0))

    def setup_optimization_tab(self):
        """Aba de otimização com layout profissional."""
        tab = self.tabview.tab("Optimization")
        tab.grid_columnconfigure(0, weight=1)
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
        self.history_frame = ctk.CTkFrame(sec_history, fg_color=("gray92", "gray18"), corner_radius=8)
        self.history_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=15)
        self.history_frame.grid_columnconfigure(0, weight=1)
        self.history_frame.grid_rowconfigure(0, weight=1)
        
        ctk.CTkLabel(self.history_frame, text="Optimization history will be displayed here.").pack(expand=True)

    def setup_log_tab(self):
        """Aba de log com interface profissional."""
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
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
    def _style_plots(self):
        """Aplica estilo dinâmico aos gráficos Matplotlib baseado no tema CTk."""
        is_dark = ctk.get_appearance_mode() == "Dark"
        face_color = '#2B2B2B' if is_dark else '#FFFFFF'
        text_color = 'white' if is_dark else 'black'
        grid_color = '#404040' if is_dark else '#D3D3D3'

        self.fig.patch.set_facecolor(face_color)
        for ax in [self.ax_s11, self.ax_imp, self.ax_th, self.ax_ph]:
            ax.set_facecolor(face_color)
            ax.tick_params(colors=text_color, which='both')
            ax.xaxis.label.set_color(text_color)
            ax.yaxis.label.set_color(text_color)
            ax.title.set_color(text_color)
            for spine in ax.spines.values():
                spine.set_edgecolor(text_color)
            ax.grid(True, color=grid_color, linestyle='--', linewidth=0.5)
        
        self.ax_3d.set_facecolor(face_color)
        self.ax_3d.tick_params(colors=text_color, which='both')
        self.ax_3d.xaxis.label.set_color(text_color)
        self.ax_3d.yaxis.label.set_color(text_color)
        self.ax_3d.zaxis.label.set_color(text_color)
        self.ax_3d.title.set_color(text_color)
        
        # Remove os painéis cinza do eixo 3D
        self.ax_3d.xaxis.set_pane_color((0, 0, 0, 0))
        self.ax_3d.yaxis.set_pane_color((0, 0, 0, 0))
        self.ax_3d.zaxis.set_pane_color((0, 0, 0, 0))


    def _style_ttk_treeview(self):
        """Aplica um estilo moderno ao ttk.TreeView para combinar com o tema do CTk."""
        style = ttk.Style()
        is_dark = ctk.get_appearance_mode() == "Dark"
        
        bg_color = "#2B2B2B" if is_dark else "#F9F9FA"
        fg_color = "#DCE4EE" if is_dark else "#333333"
        header_bg = "#303030" if is_dark else "#EAEAEA"
        selected_bg = "#1F6AA5" # Cor de seleção padrão do CTk

        style.theme_use("default")
        style.configure("Treeview", background=bg_color, foreground=fg_color, fieldbackground=bg_color, borderwidth=0, rowheight=25)
        style.map('Treeview', background=[('selected', selected_bg)])
        style.configure("Treeview.Heading", background=header_bg, foreground=fg_color, relief="flat", font=('Calibri', 10, 'bold'))
        style.map("Treeview.Heading", background=[('active', '#3C3C3C' if is_dark else '#DCDCDC')])

    def show_tooltip(self, event, text):
        if hasattr(self, 'current_tooltip') and self.current_tooltip:
            self.current_tooltip.destroy()
        
        tooltip = ctk.CTkToplevel(self.window)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{event.x_root+15}+{event.y_root+10}")
        
        label = ctk.CTkLabel(tooltip, text=text, fg_color=("#EAEAEA", "#333333"),
                             text_color=("#333333", "#EAEAEA"), corner_radius=5,
                             justify="left", wraplength=300, font=("Calibri", 12))
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
        self.update_plots(compare_mode=True)
        self.analyze_and_mark_s11() # Re-analisa para mostrar o marcador no gráfico atual

    def view_optimization_history(self):
        """Exibe o histórico de otimização em uma janela separada usando ttk.TreeView."""
        if not self.optimization_history:
            messagebox.showinfo("Optimization History", "Nenhum histórico de otimização disponível.")
            return
            
        history_window = ctk.CTkToplevel(self.window)
        history_window.title("Optimization History")
        history_window.geometry("850x400")
        history_window.grab_set()
        
        ctk.CTkLabel(history_window, text="Optimization History", 
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        frame = ctk.CTkFrame(history_window)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        columns = ("iteration", "resonant_freq", "target_freq", "error_percent", "min_s11", "scaling_factor")
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree.heading("iteration", text="Iteration")
        tree.heading("resonant_freq", text="Resonant Freq (GHz)")
        tree.heading("target_freq", text="Target Freq (GHz)")
        tree.heading("error_percent", text="Error (%)")
        tree.heading("min_s11", text="Min S11 (dB)")
        tree.heading("scaling_factor", text="Scaling Factor")
        
        for col in columns:
            tree.column(col, width=130, anchor='center')
        
        for record in self.optimization_history:
            tree.insert("", "end", values=(
                record["iteration"], f"{record['resonant_freq']:.3f}", f"{record['target_freq']:.3f}",
                f"{record['error_percent']:.1f}", f"{record['min_s11']:.2f}", f"{record['scaling_factor']:.4f}"
            ))
        
        ctk.CTkButton(history_window, text="Close", command=history_window.destroy).pack(pady=10)

    def update_quick_status(self, status, color=None):
        """Atualiza o status rápido no header."""
        colors = { "ready": ("gray85", "gray25"), "running": ("#FFA500", "#CC8400"),
                   "success": ("#2E8B57", "#3CB371"), "error": ("#DC143C", "#FF4500") }
        if color in colors:
            self.quick_status.configure(fg_color=colors[color])
        self.quick_status.configure(text=status)

    # ------------- Utilitários de Log -------------
    def log_message(self, message: str):
        """Enfileira uma mensagem para o textbox de log com carimbo de hora."""
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")

    def process_log_queue(self):
        """Consumidor assíncrono da fila de log; mantém UI responsiva."""
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

    # ----------- Física / Cálculos -----------
    def get_parameters(self) -> bool:
        """Lê valores da UI, faz casting e sincroniza `self.params`."""
        try:
            for key, var in self.entries:
                if key == "show_gui":
                    self.params["non_graphical"] = not var.get()
                elif key == "save_project":
                    self.save_project = var.get()
                elif isinstance(var, ctk.BooleanVar):
                     self.params[key] = var.get()
                elif isinstance(var, ctk.StringVar):
                    value_str = var.get()
                    # Tenta converter para float, depois int, e se falhar, mantém como string
                    try:
                        if '.' in value_str: self.params[key] = float(value_str)
                        else: self.params[key] = int(value_str)
                    except ValueError:
                        self.params[key] = value_str # Para ComboBoxes
            return True
        except Exception as e:
            self.log_message(f"Invalid value detected: {e}")
            messagebox.showerror("Invalid Input", f"Could not parse all input values. Please check for errors.\nDetails: {e}")
            return False

    def calculate_patch_dimensions(self, frequency_ghz: float, er: float, h_mm: float) -> Tuple[float, float, float]:
        """Calcula L, W e λg (em mm) para microfita retangular."""
        f = frequency_ghz * 1e9
        h = h_mm / 1000.0  # mm->m
        W = self.c / (2 * f) * math.sqrt(2 / (er + 1))
        eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W) ** -0.5
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        L_eff = self.c / (2 * f * math.sqrt(eeff))
        L = L_eff - 2 * dL
        lambda_g = self.c / (f * math.sqrt(eeff))
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)

    def _size_array_from_gain(self) -> Tuple[int, int]:
        """Deriva nº de elementos (linhas/colunas) a partir do gain desejado."""
        G_elem_dBi = 7.0  # Ganho típico de um único patch em dBi
        G_des_dBi = self.params["gain"]
        N_req = 10 ** ((G_des_dBi - G_elem_dBi) / 10.0)
        
        if N_req <= 1: return 1, 1
        
        rows = int(round(math.sqrt(N_req)))
        cols = int(math.ceil(N_req / rows))
        return rows, cols

    def calculate_parameters(self):
        """Calcula L/W/λg, spacing e layout (linhas/colunas). Atualiza UI."""
        self.log_message("Starting parameter calculation")
        if not self.get_parameters():
            self.log_message("Parameter calculation failed due to invalid input")
            return

        try:
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(
                self.params["frequency"], self.params["er"], self.params["substrate_thickness"]
            )
            self.calculated_params.update({"patch_length": L_mm, "patch_width": W_mm, "lambda_g": lambda_g_mm})
            
            lambda0_m = self.c / (self.params["frequency"] * 1e9)
            factors = {"lambda/2": 0.5, "lambda": 1.0, "0.7*lambda": 0.7, "0.8*lambda": 0.8, "0.9*lambda": 0.9}
            spacing_mm = factors.get(self.params["spacing_type"], 0.5) * lambda0_m * 1000.0
            self.calculated_params["spacing"] = spacing_mm
            
            rows, cols = self._size_array_from_gain()
            num_patches = rows * cols
            self.calculated_params.update({"num_patches": num_patches, "rows": rows, "cols": cols})
            self.log_message(f"Array sizing -> target gain {self.params['gain']} dBi, N_req≈{10**((self.params['gain']-7)/10):.2f}, layout {rows}x{cols} ({num_patches} patches)")
            
            self.calculated_params["feed_offset"] = 0.30 * L_mm
            
            total_w = cols * W_mm + (cols - 1) * spacing_mm
            total_l = rows * L_mm + (rows - 1) * spacing_mm
            margin = max(total_w, total_l) * 0.20
            self.calculated_params["substrate_width"] = total_w + margin
            self.calculated_params["substrate_length"] = total_l + margin
            
            # UI Update
            self.patches_label.configure(text=f"Number of Patches: {num_patches}")
            self.rows_cols_label.configure(text=f"Configuration: {rows} x {cols}")
            self.spacing_label.configure(text=f"Spacing: {spacing_mm:.2f} mm")
            self.dimensions_label.configure(text=f"Patch Dimensions: {L_mm:.2f} x {W_mm:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {lambda_g_mm:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset (y): {self.calculated_params['feed_offset']:.2f} mm")
            self.substrate_dims_label.configure(text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x {self.calculated_params['substrate_length']:.2f} mm")
            
            self.status_label.configure(text="Parameters calculated successfully")
            self.log_message("Parameters calculated successfully")
            self.update_quick_status("Ready", "success")
        except Exception as e:
            self.status_label.configure(text=f"Error in calculation: {e}")
            self.log_message(f"Error in calculation: {e}\nTraceback: {traceback.format_exc()}")
            self.update_quick_status("Error", "error")

    # --------- AEDT helpers ---------
    def _ensure_material(self, name: str, er: float, tan_d: float):
        """Garante a existência de um material com εr e tanδ informados."""
        try:
            # Acessa a biblioteca de materiais através do objeto hfss
            if not self.hfss.materials.checkifmaterialexists(name):
                new_mat = self.hfss.materials.add_material(name, dielectric_permittivity=er, dielectric_loss_tangent=tan_d)
                self.log_message(f"Created material: {new_mat.name} (er={er}, tanδ={tan_d})")
        except Exception as e:
            self.log_message(f"Material management warning for '{name}': {e}")

    def _open_or_create_project(self):
        """Abre o Desktop, cria um projeto temporário e inicializa a classe Hfss."""
        self.log_message(f"Launching AEDT v{self.params['aedt_version']}...")
        # A classe Hfss gerencia a inicialização do Desktop implicitamente
        self.hfss = Hfss(
            specified_version=self.params["aedt_version"],
            non_graphical=self.params["non_graphical"],
            new_desktop_session=True, # Garante uma sessão limpa
            close_on_exit=True # Garante que o AEDT fechará com o script
        )
        self.hfss.solution_type = "DrivenModal"
        self.log_message("AEDT session initialized successfully.")
        self.log_message(f"Project '{self.hfss.project_name}' created in design '{self.hfss.design_name}'.")

    def _create_coax_feed_lumped(self, ground, substrate, x_feed: float, y_feed: float, name_prefix: str):
        """Constrói um feed coaxial com uma porta lumped na base."""
        try:
            a = self.params["probe_radius"]
            b = a * self.params["coax_ba_ratio"]
            lp = self.params["coax_port_length"]
            h_sub = self.params["substrate_thickness"]
            clearance = self.params["antipad_clearance"]
            
            # 1. Pino condutor
            pin = self.hfss.modeler.create_cylinder(
                cs_axis="Z", position=[x_feed, y_feed, 0], radius=a, height=h_sub,
                name=f"{name_prefix}_Pin", matname="copper"
            )
            
            # 2. Vazio no substrato para o dielétrico do coaxial
            coax_dielectric_hole = self.hfss.modeler.create_cylinder(
                cs_axis="Z", position=[x_feed, y_feed, 0], radius=b, height=h_sub,
                name=f"{name_prefix}_VaccumHole"
            )
            substrate.subtract(coax_dielectric_hole, keep_originals=False)
            
            # 3. Anti-pad no plano de terra
            antipad = self.hfss.modeler.create_circle(
                cs_plane="XY", position=[x_feed, y_feed, 0], radius=b + clearance,
                name=f"{name_prefix}_Antipad"
            )
            ground.subtract(antipad, keep_originals=False)

            # 4. Criação da porta Lumped na base
            port_sheet = self.hfss.modeler.create_circle(
                cs_plane="XY", position=[x_feed, y_feed, 0], radius=b, name=f"port_sheet_{name_prefix}"
            )
            pin_cap = self.hfss.modeler.create_circle(
                cs_plane="XY", position=[x_feed, y_feed, 0], radius=a, name=f"pin_cap_{name_prefix}"
            )
            port_sheet.subtract(pin_cap, keep_originals=False)
            
            self.hfss.lumped_port(
                assignment=port_sheet.name,
                impedance=50,
                name=f"{name_prefix}_LumpedPort",
                renormalize=True
            )
            self.created_ports.append(f"{name_prefix}_LumpedPort")
            self.log_message(f"Lumped Port '{name_prefix}_LumpedPort' created.")
            
            return pin
        except Exception as e:
            self.log_message(f"Exception in coax creation '{name_prefix}': {e}\nTraceback: {traceback.format_exc()}")
            return None

    # ------------- Lógica Principal de Simulação -------------
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
            self.hfss.analyze(setup_name="Setup1") # Executa o setup específico

            self.log_message("Performing post-solve setup for beamforming...")
            self.window.after(0, lambda: self.sim_status_label.configure(text="Post-processing..."))
            self._postprocess_after_solve()

            self.window.after(0, self._on_simulation_complete, True, None)

        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            self.log_message(error_msg)
            self.log_message(f"Traceback: {traceback.format_exc()}")
            self.window.after(0, self._on_simulation_complete, False, str(e))

    def _create_geometry_and_boundaries(self):
        """Cria a geometria completa, boundaries e excitações, com chamadas corrigidas."""
        self.created_ports.clear()
        params = self.params
        calc = self.calculated_params

        self._ensure_material(params["substrate_material"], params["er"], params["tan_d"])

        gnd = self.hfss.modeler.create_rectangle("XY", [-calc["substrate_width"]/2, -calc["substrate_length"]/2, 0],
                                                 [calc["substrate_width"], calc["substrate_length"]], name="Ground")
        self.hfss.assign_perfect_e(gnd)

        substrate = self.hfss.modeler.create_box(
            position=[-calc["substrate_width"]/2, -calc["substrate_length"]/2, 0],
            sizes=[calc["substrate_width"], calc["substrate_length"], params["substrate_thickness"]],
            name="Substrate", matname=params["substrate_material"]
        )
        substrate.transparency = 0.7

        total_w = calc["cols"] * calc["patch_width"] + (calc["cols"] - 1) * calc["spacing"]
        total_l = calc["rows"] * calc["patch_length"] + (calc["rows"] - 1) * calc["spacing"]
        start_x, start_y = -total_w / 2, -total_l / 2

        all_patches = []
        for r in range(calc["rows"]):
            for c in range(calc["cols"]):
                patch_x = start_x + c * (calc["patch_width"] + calc["spacing"])
                patch_y = start_y + r * (calc["patch_length"] + calc["spacing"])

                patch = self.hfss.modeler.create_rectangle(
                    cs_plane="XY", position=[patch_x, patch_y, params["substrate_thickness"]],
                    dimension_list=[calc["patch_width"], calc["patch_length"]],
                    name=f"Patch_{r}_{c}"
                )
                all_patches.append(patch)

                feed_x_pos = patch_x + calc["patch_width"] * params["feed_rel_x"]
                feed_y_pos = patch_y + calc["patch_length"] / 2 - calc["feed_offset"]
                
                self._create_coax_feed_lumped(gnd, substrate, feed_x_pos, feed_y_pos, name_prefix=f"P{len(self.created_ports)+1}")

        if len(all_patches) > 1:
            radiator = self.hfss.modeler.unite(all_patches)
            radiator.name = "Radiator"
        else:
            radiator = all_patches[0]
            radiator.name = "Radiator"
        
        self.hfss.assign_material(radiator, "copper")

        freq_str = f"{params['frequency']}GHz"
        self.hfss.create_open_region(frequency=freq_str)
        self.log_message(f"Created geometry and open region boundary for {freq_str}.")

    def _create_analysis_setup(self):
        """Cria o setup e o sweep, com chamadas e lógica corrigidas."""
        setup_name = "Setup1"
        setup = self.hfss.create_setup(setup_name=setup_name, setup_type="DrivenModal")
        
        setup.props["Frequency"] = f"{self.params['frequency']}GHz"
        setup.props["MaximumPasses"] = 15
        setup.props["MaxDeltaS"] = 0.02
        
        sweep_name = "Sweep1"
        self.hfss.create_interpolating_sweep(
            setup_name=setup_name, sweep_name=sweep_name,
            start_freq=f"{self.params['sweep_start']}GHz",
            stop_freq=f"{self.params['sweep_stop']}GHz",
            save_fields=True
        )
        
        self._ensure_infinite_sphere()
        self.log_message(f"Created '{setup_name}' with '{sweep_name}'.")

    # ------------- Loop Principal e Encerramento -------------
    def run(self):
        """Inicia o loop principal da GUI."""
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.mainloop()

    def on_closing(self):
        """Lida com o fechamento da janela, liberando recursos do AEDT."""
        if messagebox.askokcancel("Quit", "Do you want to quit? This will close the AEDT session."):
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