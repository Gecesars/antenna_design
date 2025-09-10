# -*- coding: utf-8 -*-
"""
Modern Patch Antenna Designer (v4.2) - Versão Corrigida
-------------------------------------------------------
Correções implementadas:
- Correção completa dos erros de criação de variáveis de pós-processamento
- Separação de abas para S11/Impedância e diagramas de radiação
- Melhor tratamento de exceções em operações críticas
- Correção de problemas com nomes de objetos no AEDT
- Melhor gerenciamento de recursos e memória
"""

import os
import re
import tempfile
import time
import threading
from datetime import datetime
import math
import json
import traceback
import queue
from typing import Tuple, List, Optional, Dict, Any

import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.mplot3d import Axes3D
import customtkinter as ctk
from tkinter import messagebox

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
        self.simulation_thread = None

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
            "frequency": 10.0,
            "gain": 12.0,
            "sweep_start": 8.0,
            "sweep_stop": 12.0,
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,
            "spacing_type": "lambda/2",
            "substrate_material": "Duroid (tm)",
            "substrate_thickness": 0.5,
            "metal_thickness": 0.035,
            "er": 2.2,
            "tan_d": 0.0009,
            "feed_position": "inset",
            "feed_rel_x": 0.485,
            "probe_radius": 0.40,
            "coax_ba_ratio": 2.3,
            "coax_wall_thickness": 0.20,
            "coax_port_length": 3.0,
            "antipad_clearance": 0.10,
            "sweep_type": "Interpolating",
            "sweep_step": 0.02,
            "theta_step": 10.0,
            "phi_step": 10.0
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
        """Constroi a janela principal e abas com layout profissional."""
        self.window = ctk.CTk()
        self.window.title("Modern Patch Antenna Array Designer")
        self.window.geometry("1600x1000")
        
        # Configurar grid principal
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        # Header profissional
        header = ctk.CTkFrame(self.window, height=80, fg_color=("gray90", "gray15"))
        header.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 5))
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)
        
        # Logo placeholder
        logo_frame = ctk.CTkFrame(header, width=60, height=60, fg_color=("gray80", "gray25"))
        logo_frame.grid(row=0, column=0, padx=15, pady=10, sticky="w")
        logo_frame.grid_propagate(False)
        ctk.CTkLabel(logo_frame, text="ANT", font=ctk.CTkFont(size=20, weight="bold")).pack(expand=True)
        
        # Título principal
        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkLabel(
            title_frame, text="Modern Patch Antenna Array Designer",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=("gray10", "gray90")
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_frame, text="Professional RF Design Tool",
            font=ctk.CTkFont(size=14),
            text_color=("gray40", "gray60")
        ).pack(anchor="w")

        # Status rápido
        status_frame = ctk.CTkFrame(header, fg_color="transparent")
        status_frame.grid(row=0, column=2, padx=15, pady=10, sticky="e")
        self.quick_status = ctk.CTkLabel(status_frame, text="Ready", font=ctk.CTkFont(weight="bold"), 
                                        fg_color=("gray85", "gray25"), corner_radius=8)
        self.quick_status.pack(padx=5, pady=5, ipadx=10, ipady=5)

        # Sistema de abas principal
        self.tabview = ctk.CTkTabview(self.window, fg_color=("gray95", "gray10"))
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Abas
        tabs = ["Design", "Simulation", "Results", "Optimization", "Log"]
        for name in tabs:
            self.tabview.add(name)
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        self.setup_design_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_optimization_tab()
        self.setup_log_tab()

        # Barra de status inferior
        status = ctk.CTkFrame(self.window, height=40, fg_color=("gray92", "gray18"))
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        status.grid_columnconfigure(0, weight=1)
        
        self.status_label = ctk.CTkLabel(status, text="Ready", font=ctk.CTkFont(weight="bold"), 
                                        anchor="w", text_color=("gray30", "gray70"))
        self.status_label.grid(row=0, column=0, padx=15, pady=6, sticky="w")
        
        version_label = ctk.CTkLabel(status, text="v4.2 © 2024 RF Design Suite", 
                                    font=ctk.CTkFont(size=12), text_color=("gray40", "gray60"))
        version_label.grid(row=0, column=1, padx=15, pady=6, sticky="e")

        self.process_log_queue()

    def create_section(self, parent, title, description=None, row=0, column=0, padx=10, pady=10, colspan=1):
        """Cria um frame com título and descrição para organizar a UI."""
        section = ctk.CTkFrame(parent, fg_color=("gray97", "gray12"), corner_radius=8)
        section.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady, columnspan=colspan)
        section.grid_columnconfigure(0, weight=1)
        
        # Header da seção
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
        
        # Container principal com scroll
        main = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=1)

        # Seção de parâmetros de antena
        sec_ant = self.create_section(main, "Antenna Parameters", "Fundamental design parameters", 0, 0)
        self.entries = []
        row_idx = 2

        def add_entry(section, label, key, value, row, tooltip=None, combo=None, check=False, unit=""):
            # Frame para cada campo
            field_frame = ctk.CTkFrame(section, fg_color="transparent")
            field_frame.grid(row=row, column=0, sticky="ew", padx=15, pady=4)
            field_frame.grid_columnconfigure(0, weight=3)
            field_frame.grid_columnconfigure(1, weight=5)
            
            # Label
            label_text = f"{label}:"
            if unit:
                label_text = f"{label} ({unit}):"
                
            lbl = ctk.CTkLabel(field_frame, text=label_text, font=ctk.CTkFont(weight="bold"), 
                              anchor="w", text_color=("gray30", "gray70"))
            lbl.grid(row=0, column=0, sticky="w", padx=(0, 10))
            
            # Tooltip
            if tooltip:
                info_btn = ctk.CTkLabel(field_frame, text="ⓘ", font=ctk.CTkFont(size=10), 
                                       text_color=("gray50", "gray50"), cursor="hand2")
                info_btn.grid(row=0, column=1, sticky="e", padx=(0, 5))
                info_btn.bind("<Enter>", lambda e, t=tooltip: self.show_tooltip(e, t))
                info_btn.bind("<Leave>", self.hide_tooltip)
            
            # Widget de entrada
            input_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
            input_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))
            input_frame.grid_columnconfigure(0, weight=1)
            
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(input_frame, values=combo, variable=var, 
                                        font=ctk.CTkFont(size=13), dropdown_font=ctk.CTkFont(size=12),
                                        height=32, corner_radius=6)
                widget.grid(row=0, column=0, sticky="ew")
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(input_frame, text="", variable=var, 
                                        width=20, height=20, corner_radius=4)
                widget.grid(row=0, column=0, sticky="w")
                self.entries.append((key, var))
            else:
                widget = ctk.CTkEntry(input_frame, font=ctk.CTkFont(size=13),
                                     height=32, corner_radius=6)
                widget.insert(0, str(value))
                widget.grid(row=0, column=0, sticky="ew")
                self.entries.append((key, widget))
            
            return row + 1

        # Campos de parâmetros de antena
        row_idx = add_entry(sec_ant, "Central Frequency", "frequency", self.params["frequency"], row_idx, 
                           tooltip="Operating frequency of the antenna", unit="GHz")
        row_idx = add_entry(sec_ant, "Desired Gain", "gain", self.params["gain"], row_idx,
                           tooltip="Target gain in dBi", unit="dBi")
        row_idx = add_entry(sec_ant, "Sweep Start", "sweep_start", self.params["sweep_start"], row_idx,
                           tooltip="Start frequency for simulation sweep", unit="GHz")
        row_idx = add_entry(sec_ant, "Sweep Stop", "sweep_stop", self.params["sweep_stop"], row_idx,
                           tooltip="Stop frequency for simulation sweep", unit="GHz")
        row_idx = add_entry(sec_ant, "Patch Spacing", "spacing_type", self.params["spacing_type"], row_idx,
                           tooltip="Distance between patch elements", 
                           combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        # Seção de parâmetros de substrato
        sec_sub = self.create_section(main, "Substrate Parameters", "Material properties", 0, 1)
        row_idx = 2
        row_idx = add_entry(sec_sub, "Substrate Material", "substrate_material",
                            self.params["substrate_material"], row_idx,
                            tooltip="Dielectric material properties",
                            combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        row_idx = add_entry(sec_sub, "Relative Permittivity", "er", self.params["er"], row_idx,
                           tooltip="Dielectric constant (εr)", unit="")
        row_idx = add_entry(sec_sub, "Loss Tangent", "tan_d", self.params["tan_d"], row_idx,
                           tooltip="Dissipation factor (tan δ)", unit="")
        row_idx = add_entry(sec_sub, "Substrate Thickness", "substrate_thickness", 
                           self.params["substrate_thickness"], row_idx,
                           tooltip="Height of substrate material", unit="mm")
        row_idx = add_entry(sec_sub, "Metal Thickness", "metal_thickness", self.params["metal_thickness"], row_idx,
                           tooltip="Conductor thickness", unit="mm")

        # Seção de parâmetros de alimentação
        sec_coax = self.create_section(main, "Feed Parameters", "Coaxial feed configuration", 1, 0)
        row_idx = 2
        row_idx = add_entry(sec_coax, "Feed position type", "feed_position", self.params["feed_position"], row_idx,
                            tooltip="Feed connection method",
                            combo=["inset", "edge"])
        row_idx = add_entry(sec_coax, "Feed relative X position", "feed_rel_x", self.params["feed_rel_x"], row_idx,
                           tooltip="Normalized position along patch width (0-1)", unit="")
        row_idx = add_entry(sec_coax, "Inner radius", "probe_radius", self.params["probe_radius"], row_idx,
                           tooltip="Inner conductor radius", unit="mm")
        row_idx = add_entry(sec_coax, "b/a ratio", "coax_ba_ratio", self.params["coax_ba_ratio"], row_idx,
                           tooltip="Outer to inner conductor ratio", unit="")
        row_idx = add_entry(sec_coax, "Shield wall thickness", "coax_wall_thickness", 
                           self.params["coax_wall_thickness"], row_idx,
                           tooltip="Outer conductor thickness", unit="mm")
        row_idx = add_entry(sec_coax, "Port length below GND", "coax_port_length", 
                           self.params["coax_port_length"], row_idx,
                           tooltip="Port extension below ground", unit="mm")
        row_idx = add_entry(sec_coax, "Anti-pad clearance", "antipad_clearance", 
                           self.params["antipad_clearance"], row_idx,
                           tooltip="Clearance around feed", unit="mm")

        # Seção de configuração de simulação
        sec_sim = self.create_section(main, "Simulation Settings", "Solution configuration", 1, 1)
        row_idx = 2
        row_idx = add_entry(sec_sim, "CPU Cores", "cores", self.params["cores"], row_idx,
                           tooltip="Number of processing cores to use", unit="")
        row_idx = add_entry(sec_sim, "Show HFSS Interface", "show_gui", not self.params["non_graphical"], row_idx, 
                           check=True, tooltip="Display HFSS interface during simulation")
        row_idx = add_entry(sec_sim, "Save Project", "save_project", self.save_project, row_idx, check=True,
                           tooltip="Save project after simulation")
        row_idx = add_entry(sec_sim, "Sweep Type", "sweep_type", self.params["sweep_type"], row_idx,
                           tooltip="Frequency sweep method",
                           combo=["Discrete", "Interpolating", "Fast"])
        row_idx = add_entry(sec_sim, "Discrete Step", "sweep_step", self.params["sweep_step"], row_idx,
                           tooltip="Frequency step size for discrete sweep", unit="GHz")
        row_idx = add_entry(sec_sim, "3D Theta step", "theta_step", self.params["theta_step"], row_idx,
                           tooltip="Angular resolution for theta", unit="deg")
        row_idx = add_entry(sec_sim, "3D Phi step", "phi_step", self.params["phi_step"], row_idx,
                           tooltip="Angular resolution for phi", unit="deg")

        # Seção de parâmetros calculados
        sec_calc = self.create_section(main, "Calculated Parameters", "Derived design values", 2, 0, colspan=2)
        calc_grid = ctk.CTkFrame(sec_calc, fg_color="transparent")
        calc_grid.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        calc_grid.grid_columnconfigure(0, weight=1)
        calc_grid.grid_columnconfigure(1, weight=1)
        
        # Configuração de grid para os parâmetros calculados
        self.patches_label = ctk.CTkLabel(calc_grid, text="Number of Patches: 4", 
                                         font=ctk.CTkFont(weight="bold"), anchor="w")
        self.patches_label.grid(row=0, column=0, sticky="w", pady=6)
        
        self.rows_cols_label = ctk.CTkLabel(calc_grid, text="Configuration: 2 x 2", 
                                           font=ctk.CTkFont(weight="bold"), anchor="w")
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=6)
        
        self.spacing_label = ctk.CTkLabel(calc_grid, text="Spacing: -- mm", 
                                         font=ctk.CTkFont(weight="bold"), anchor="w")
        self.spacing_label.grid(row=1, column=0, sticky="w", pady=6)
        
        self.dimensions_label = ctk.CTkLabel(calc_grid, text="Patch Dimensions: -- x -- mm", 
                                            font=ctk.CTkFont(weight="bold"), anchor="w")
        self.dimensions_label.grid(row=1, column=1, sticky="w", pady=6)
        
        self.lambda_label = ctk.CTkLabel(calc_grid, text="Guided Wavelength: -- mm", 
                                        font=ctk.CTkFont(weight="bold"), anchor="w")
        self.lambda_label.grid(row=2, column=0, sticky="w", pady=6)
        
        self.feed_offset_label = ctk.CTkLabel(calc_grid, text="Feed Offset (y): -- mm", 
                                             font=ctk.CTkFont(weight="bold"), anchor="w")
        self.feed_offset_label.grid(row=2, column=1, sticky="w", pady=6)
        
        self.substrate_dims_label = ctk.CTkLabel(calc_grid, text="Substrate Dimensions: -- x -- mm",
                                                 font=ctk.CTkFont(weight="bold"), anchor="w")
        self.substrate_dims_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=6)

        # Botões de ação
        btn_frame = ctk.CTkFrame(sec_calc, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=12)
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_columnconfigure(2, weight=1)
        
        ctk.CTkButton(btn_frame, text="Calculate Parameters", command=self.calculate_parameters,
                      fg_color="#2E8B57", hover_color="#3CB371", height=36, 
                      font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=8, sticky="ew")
        
        ctk.CTkButton(btn_frame, text="Save Parameters", command=self.save_parameters,
                      fg_color="#4169E1", hover_color="#6495ED", height=36,
                      font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=8, sticky="ew")
        
        ctk.CTkButton(btn_frame, text="Load Parameters", command=self.load_parameters,
                      fg_color="#FF8C00", hover_color="#FFA500", height=36,
                      font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=8, sticky="ew")

    def setup_simulation_tab(self):
        """Aba de simulação com layout profissional."""
        tab = self.tabview.tab("Simulation")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        
        # Seção de controle de simulação
        sec_control = self.create_section(main, "Simulation Control", "Run and monitor simulations", 0, 0)
        
        ctk.CTkLabel(sec_control, text="Manage simulation execution and progress", 
                    font=ctk.CTkFont(size=13), text_color=("gray40", "gray60")).grid(row=2, column=0, sticky="w", padx=15, pady=(0, 15))
        
        # Botões de controle
        btn_frame = ctk.CTkFrame(sec_control, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=10)
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_columnconfigure(2, weight=1)
        
        self.run_button = ctk.CTkButton(btn_frame, text="▶ Run Simulation", command=self.run_simulation,
                                        fg_color="#2E8B57", hover_color="#3CB371", height=40, 
                                        font=ctk.CTkFont(size=14, weight="bold"))
        self.run_button.grid(row=0, column=0, padx=8, sticky="ew")
        
        ctk.CTkButton(btn_frame, text="⏾ Save Project", command=self.save_project_toggle,
                      fg_color="#4169E1", hover_color="#6495ED", height=40,
                      font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=2, padx=8, sticky="ew")
        
        # Status da simulação
        status_frame = ctk.CTkFrame(sec_control, fg_color=("gray92", "gray18"), corner_radius=8)
        status_frame.grid(row=4, column=0, sticky="ew", padx=15, pady=10)
        status_frame.grid_columnconfigure(0, weight=1)
        
        self.sim_status_label = ctk.CTkLabel(status_frame, text="Simulation not started", 
                                            font=ctk.CTkFont(weight="bold"), text_color=("gray30", "gray70"))
        self.sim_status_label.grid(row=0, column=0, padx=15, pady=12)
        
        # Dica
        tip_frame = ctk.CTkFrame(sec_control, fg_color=("gray94", "gray16"), corner_radius=8)
        tip_frame.grid(row=5, column=0, sticky="ew", padx=15, pady=(5, 15))
        tip_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(tip_frame, text="💡 Tip: With post variables (p_i / ph_i) you can retune beams without re-solving.",
                     font=ctk.CTkFont(size=12), text_color=("gray40", "gray60"), 
                     justify="left").grid(row=0, column=0, padx=12, pady=10)

    def setup_results_tab(self):
        """Aba de resultados com visualização profissional."""
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Criar um tabview interno para separar S11/Impedância e Diagramas de Radiação
        inner_tabview = ctk.CTkTabview(tab, fg_color=("gray95", "gray10"))
        inner_tabview.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        inner_tabview.add("S11 & Impedance")
        inner_tabview.add("Radiation Patterns")
        inner_tabview.add("3D Pattern")
        
        # Configurar cada aba interna
        self.setup_s11_tab(inner_tabview.tab("S11 & Impedance"))
        self.setup_radiation_tab(inner_tabview.tab("Radiation Patterns"))
        self.setup_3d_tab(inner_tabview.tab("3D Pattern"))

    def setup_s11_tab(self, tab):
        """Aba para gráficos de S11 e Impedância."""
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        
        # Header
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(header, text="S-Parameters & Impedance", 
                    font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="w")
        
        ctk.CTkLabel(header, text="Analyze reflection coefficient and input impedance", 
                    font=ctk.CTkFont(size=13), text_color=("gray40", "gray60")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        # Área de gráficos
        graph_frame = ctk.CTkFrame(main, fg_color=("gray96", "gray14"), corner_radius=10)
        graph_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)
        
        # Configuração da figura com tema escuro/claro
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        text_color = 'white' if ctk.get_appearance_mode() == "Dark" else 'black'
        grid_color = 'gray' if ctk.get_appearance_mode() == "Dark" else 'lightgray'
        
        self.fig_s11 = plt.figure(figsize=(12, 8), facecolor=face)
        self.fig_s11.patch.set_facecolor(face)
        
        # Subplots para S11 e Impedância
        self.ax_s11 = self.fig_s11.add_subplot(211)
        self.ax_imp = self.fig_s11.add_subplot(212)
        
        # Configurar estilo dos gráficos
        for ax in (self.ax_s11, self.ax_imp):
            ax.set_facecolor(face)
            ax.tick_params(colors=text_color)
            ax.xaxis.label.set_color(text_color)
            ax.yaxis.label.set_color(text_color)
            ax.title.set_color(text_color)
            for s in ['bottom', 'top', 'right', 'left']: 
                ax.spines[s].set_color(text_color)
            ax.grid(True, color=grid_color, alpha=0.3)
        
        # Títulos
        self.ax_s11.set_title("S-Parameter (S11)")
        self.ax_imp.set_title("Input Impedance")
        
        # Canvas
        self.canvas_s11 = FigureCanvasTkAgg(self.fig_s11, master=graph_frame)
        self.canvas_s11.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        
        # Painel de controle
        control_frame = ctk.CTkFrame(main, fg_color="transparent")
        control_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        control_frame.grid_columnconfigure(0, weight=1)
        
        # Controles de resultado
        btn_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=0, sticky="ew", pady=5)
        
        ctk.CTkButton(btn_frame, text="Analyze S11", command=self.analyze_and_mark_s11,
                      fg_color="#6A5ACD", hover_color="#7B68EE", width=120).pack(side="left", padx=5)
        
        ctk.CTkButton(btn_frame, text="Export PNG", command=self.export_png_s11,
                      fg_color="#20B2AA", hover_color="#40E0D0", width=120).pack(side="left", padx=5)
        
        ctk.CTkButton(btn_frame, text="Export CSV (S11)", command=self.export_csv,
                      fg_color="#6A5ACD", hover_color="#7B68EE", width=120).pack(side="left", padx=5)
        
        # Label de resultado
        self.result_label = ctk.CTkLabel(control_frame, text="", font=ctk.CTkFont(weight="bold"),
                                        text_color=("gray30", "gray70"), height=30)
        self.result_label.grid(row=1, column=0, sticky="ew", pady=(5, 0))

    def setup_radiation_tab(self, tab):
        """Aba para diagramas de radiação 2D."""
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        
        # Header
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(header, text="Radiation Patterns", 
                    font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="w")
        
        ctk.CTkLabel(header, text="Analyze antenna radiation characteristics", 
                    font=ctk.CTkFont(size=13), text_color=("gray40", "gray60")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        # Área de gráficos
        graph_frame = ctk.CTkFrame(main, fg_color=("gray96", "gray14"), corner_radius=10)
        graph_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)
        
        # Configuração da figura com tema escuro/claro
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        text_color = 'white' if ctk.get_appearance_mode() == "Dark" else 'black'
        grid_color = 'gray' if ctk.get_appearance_mode() == "Dark" else 'lightgray'
        
        self.fig_radiation = plt.figure(figsize=(12, 8), facecolor=face)
        self.fig_radiation.patch.set_facecolor(face)
        
        # Subplots para padrões de radiação
        self.ax_th = self.fig_radiation.add_subplot(211)
        self.ax_ph = self.fig_radiation.add_subplot(212)
        
        # Configurar estilo dos gráficos
        for ax in (self.ax_th, self.ax_ph):
            ax.set_facecolor(face)
            ax.tick_params(colors=text_color)
            ax.xaxis.label.set_color(text_color)
            ax.yaxis.label.set_color(text_color)
            ax.title.set_color(text_color)
            for s in ['bottom', 'top', 'right', 'left']: 
                ax.spines[s].set_color(text_color)
            ax.grid(True, color=grid_color, alpha=0.3)
        
        # Títulos
        self.ax_th.set_title("Radiation Pattern - Theta Cut (Phi=0°)")
        self.ax_ph.set_title("Radiation Pattern - Phi Cut (Theta=90°)")
        
        # Canvas
        self.canvas_radiation = FigureCanvasTkAgg(self.fig_radiation, master=graph_frame)
        self.canvas_radiation.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        
        # Painel de controle
        control_frame = ctk.CTkFrame(main, fg_color="transparent")
        control_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        control_frame.grid_columnconfigure(0, weight=1)
        
        # Controles de resultado
        btn_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=0, sticky="ew", pady=5)
        
        ctk.CTkButton(btn_frame, text="Refresh Patterns", command=self.refresh_patterns_only,
                      fg_color="#FF8C00", hover_color="#FFA500", width=140).pack(side="left", padx=5)
        
        self.auto_refresh_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(btn_frame, text="Auto-refresh (1.5s)", variable=self.auto_refresh_var,
                        command=self.toggle_auto_refresh, width=140).pack(side="left", padx=5)
        
        # Frame de controles de beamforming
        self.src_frame = ctk.CTkFrame(control_frame, fg_color=("gray94", "gray16"), corner_radius=8)
        self.src_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(5, 10))
        self.source_controls: Dict[str, Dict[str, ctk.CTkBaseClass]] = {}

    def setup_3d_tab(self, tab):
        """Aba para diagrama de radiação 3D."""
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        
        # Header
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(header, text="3D Radiation Pattern", 
                    font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="w")
        
        ctk.CTkLabel(header, text="Visualize 3D radiation characteristics", 
                    font=ctk.CTkFont(size=13), text_color=("gray40", "gray60")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        # Área de gráficos
        graph_frame = ctk.CTkFrame(main, fg_color=("gray96", "gray14"), corner_radius=10)
        graph_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)
        
        # Configuração da figura com tema escuro/claro
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        text_color = 'white' if ctk.get_appearance_mode() == "Dark" else 'black'
        
        self.fig_3d = plt.figure(figsize=(10, 8), facecolor=face)
        self.fig_3d.patch.set_facecolor(face)
        
        # Subplot 3D
        self.ax_3d = self.fig_3d.add_subplot(111, projection='3d')
        
        # Configurar estilo do gráfico 3D
        self.ax_3d.set_facecolor(face)
        self.ax_3d.tick_params(colors=text_color)
        self.ax_3d.title.set_color(text_color)
        
        # Título
        self.ax_3d.set_title("3D Radiation Pattern")
        
        # Canvas
        self.canvas_3d = FigureCanvasTkAgg(self.fig_3d, master=graph_frame)
        self.canvas_3d.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)

    def setup_optimization_tab(self):
        """Nova aba dedicada para otimização."""
        tab = self.tabview.tab("Optimization")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        
        # Header
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        header.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(header, text="Design Optimization", 
                    font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="w")
        
        ctk.CTkLabel(header, text="Automatically optimize design based on simulation results", 
                    font=ctk.CTkFont(size=13), text_color=("gray40", "gray60")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        # Seção de otimização
        sec_opt = self.create_section(main, "Optimization Control", "Tune design for better performance", 1, 0)
        
        # Status de otimização
        status_frame = ctk.CTkFrame(sec_opt, fg_color=("gray92", "gray18"), corner_radius=8)
        status_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=10)
        status_frame.grid_columnconfigure(0, weight=1)
        
        self.opt_status_label = ctk.CTkLabel(status_frame, text="No optimization performed yet", 
                                            font=ctk.CTkFont(weight="bold"), text_color=("gray30", "gray70"))
        self.opt_status_label.grid(row=0, column=0, padx=15, pady=12)
        
        # Botões de otimização
        btn_frame = ctk.CTkFrame(sec_opt, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=15)
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_columnconfigure(2, weight=1)
        
        ctk.CTkButton(btn_frame, text="Analyze & Optimize", command=self.analyze_and_optimize,
                      fg_color="#2E8B57", hover_color="#3CB371", height=40,
                      font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=8, sticky="ew")
        
        ctk.CTkButton(btn_frame, text="Reset to Original", command=self.reset_to_original,
                      fg_color="#DC143C", hover_color="#FF4500", height=40,
                      font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=8, sticky="ew")
        
        ctk.CTkButton(btn_frame, text="View History", command=self.view_optimization_history,
                      fg_color="#4169E1", hover_color="#6495ED", height=40,
                      font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=8, sticky="ew")
        
        # Seção de histórico de otimização
        sec_history = self.create_section(main, "Optimization History", "Track design changes", 2, 0)
        
        # Treeview para histórico (placeholder)
        history_frame = ctk.CTkFrame(sec_history, fg_color=("gray92", "gray18"), corner_radius=8)
        history_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=15)
        history_frame.grid_columnconfigure(0, weight=1)
        history_frame.grid_rowconfigure(0, weight=1)
        
        # Placeholder para tabela de histórico
        ctk.CTkLabel(history_frame, text="Optimization history will appear here after runs",
                    font=ctk.CTkFont(size=13), text_color=("gray40", "gray60")).pack(expand=True)

    def setup_log_tab(self):
        """Aba de log com interface profissional."""
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        
        # Header
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(header, text="Simulation Log", 
                    font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="w")
        
        ctk.CTkLabel(header, text="View detailed simulation messages and events", 
                    font=ctk.CTkFont(size=13), text_color=("gray40", "gray60")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        # Área de log
        log_frame = ctk.CTkFrame(main, fg_color=("gray96", "gray14"), corner_radius=10)
        log_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
        self.log_text = ctk.CTkTextbox(log_frame, width=900, height=500, 
                                      font=ctk.CTkFont(family="Consolas", size=12),
                                      fg_color=("gray98", "gray10"), text_color=("gray20", "gray80"))
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        
        # Botões de ação
        btn_frame = ctk.CTkFrame(main, fg_color="transparent")
        btn_frame.grid(row=2, column=0, sticky="ew", pady=10)
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_columnconfigure(2, weight=1)
        
        ctk.CTkButton(btn_frame, text="Clear Log", command=self.clear_log,
                      fg_color="#DC143C", hover_color="#FF4500").grid(row=0, column=0, padx=8)
        
        ctk.CTkButton(btn_frame, text="Save Log", command=self.save_log,
                      fg_color="#4169E1", hover_color="#6495ED").grid(row=0, column=1, padx=8)
        
        ctk.CTkButton(btn_frame, text="Export Log", command=self.export_log,
                      fg_color="#2E8B57", hover_color="#3CB371").grid(row=0, column=2, padx=8)

    # ------------- Novas funcionalidades de GUI -------------
    def show_tooltip(self, event, text):
        """Exibe um tooltip para o controle."""
        tooltip = ctk.CTkToplevel(self.window)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
        
        label = ctk.CTkLabel(tooltip, text=text, 
                            fg_color=("gray90", "gray20"), 
                            text_color=("gray20", "gray80"),
                            corner_radius=5, justify="left", wraplength=300)
        label.pack(padx=5, pady=5)
        
        # Guardar referência para fechar depois
        self.current_tooltip = tooltip
        tooltip.after(3000, tooltip.destroy)

    def hide_tooltip(self, event):
        """Esconde o tooltip atual."""
        if hasattr(self, 'current_tooltip') and self.current_tooltip:
            self.current_tooltip.destroy()

    def save_project_toggle(self):
        """Alterna o estado de salvamento do projeto."""
        self.save_project = not self.save_project
        status = "ON" if self.save_project else "OFF"
        self.log_message(f"Project saving {status}")
        self.status_label.configure(text=f"Project saving {status}")

    def export_png_s11(self):
        """Exporta a figura S11 para PNG."""
        try:
            if hasattr(self, 'fig_s11'):
                self.fig_s11.savefig("s11_results.png", dpi=300, bbox_inches='tight')
                self.log_message("S11 plot saved to s11_results.png")
        except Exception as e:
            self.log_message(f"Error saving S11 plot: {e}")

    def export_report(self):
        """Exporta um relatório completo em PDF."""
        self.log_message("Export report functionality would be implemented here")

    def compare_results(self):
        """Compara resultados entre diferentes simulações."""
        self.log_message("Compare results functionality would be implemented here")

    def view_optimization_history(self):
        """Exibe o histórico de otimização em uma janela separada."""
        if not self.optimization_history:
            messagebox.showinfo("Optimization History", "No optimization history available.")
            return
            
        history_window = ctk.CTkToplevel(self.window)
        history_window.title("Optimization History")
        history_window.geometry("800x500")
        history_window.grab_set()
        
        # Header
        ctk.CTkLabel(history_window, text="Optimization History", 
                    font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        # Treeview para histórico
        frame = ctk.CTkFrame(history_window)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Colunas
        columns = ("iteration", "resonant_freq", "target_freq", "error_percent", "min_s11", "scaling_factor")
        tree = ctk.CTkTreeview(frame, columns=columns, show="headings")
        
        # Definir cabeçalhos
        tree.heading("iteration", text="Iteration")
        tree.heading("resonant_freq", text="Resonant Freq (GHz)")
        tree.heading("target_freq", text="Target Freq (GHz)")
        tree.heading("error_percent", text="Error (%)")
        tree.heading("min_s11", text="Min S11 (dB)")
        tree.heading("scaling_factor", text="Scaling Factor")
        
        # Definir larguras
        tree.column("iteration", width=80)
        tree.column("resonant_freq", width=120)
        tree.column("target_freq", width=120)
        tree.column("error_percent", width=80)
        tree.column("min_s11", width=100)
        tree.column("scaling_factor", width=100)
        
        # Adicionar dados
        for record in self.optimization_history:
            tree.insert("", "end", values=(
                record["iteration"],
                f"{record['resonant_freq']:.3f}",
                f"{record['target_freq']:.3f}",
                f"{record['error_percent']:.1f}",
                f"{record['min_s11']:.2f}",
                f"{record['scaling_factor']:.3f}"
            ))
        
        tree.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Botão de fechar
        ctk.CTkButton(history_window, text="Close", command=history_window.destroy).pack(pady=10)

    def export_log(self):
        """Exporta o log em formato de texto."""
        try:
            filename = f"simulation_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, "w", encoding="utf-8") as f:
                f.write(self.log_text.get("1.0", "end"))
            self.log_message(f"Log exported to {filename}")
        except Exception as e:
            self.log_message(f"Error exporting log: {e}")

    # ------------- Atualização do status rápido -------------
    def update_quick_status(self, status, color=None):
        """Atualiza o status rápido no header."""
        colors = {
            "ready": ("gray85", "gray25"),
            "running": ("#FFA500", "#CC8400"),
            "success": ("#2E8B57", "#3CB371"),
            "error": ("#DC143C", "#FF4500")
        }
        
        if color and color in colors:
            self.quick_status.configure(fg_color=colors[color])
        self.quick_status.configure(text=status)

    # ------------- Utilidades de Log -------------
    def log_message(self, message: str):
        """Enfileira uma mensagem para o textbox de log com carimbo de hora."""
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")

    def process_log_queue(self):
        """Consumidor assíncrono da fila de log; mantém UI responsiva."""
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
        """Valida faixas de parâmetros mais críticas."""
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
            ok = False; msgs.append("substrate_thickness must be > 0")
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
        """Lê valores da UI, faz casting e sincroniza `self.params`."""
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
        return self._validate_ranges()

    def calculate_patch_dimensions(self, frequency_ghz: float) -> Tuple[float, float, float]:
        """Calcula L, W e λg (em mm) para microfita retangular."""
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

    def _size_array_from_gain(self) -> Tuple[int, int, int]:
        """Deriva nº de elementos (linhas/colunas) a partir do gain desejado."""
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
        """Define dimensões do substrato com margin de 20% do maior lado útil."""
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
        """Calcula L/W/λg, spacing e layout (linhas/colunas). Atualiza UI."""
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
            self.update_quick_status("Ready", "success")
        except Exception as e:
            self.status_label.configure(text=f"Error in calculation: {e}")
            self.log_message(f"Error in calculation: {e}\nTraceback: {traceback.format_exc()}")
            self.update_quick_status("Error", "error")

    # --------- AEDT helpers ---------
    def _ensure_material(self, name: str, er: float, tan_d: float):
        """Garante a existência de um material com εr and tanδ informados."""
        try:
            if not self.hfss.materials.checkifmaterialexists(name):
                self.hfss.materials.add_material(name)
                m = self.hfss.materials.material_keys[name]
                m.permittivity = er; m.dielectric_loss_tangent = tan_d
                self.log_message(f"Created material: {name} (er={er}, tanδ={tan_d})")
        except Exception as e:
            self.log_message(f"Material management warning for '{name}': {e}")

    def _open_or_create_project(self):
        """Abre Desktop/Projeto temporário e cria design HFSS DrivenModal."""
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
        """Cria/atualiza variáveis do design em HFSS (unidades mm/GHz)."""
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
        """Constrói pino, blindagem e porta lumped no plano inferior."""
        try:
            a_val = float(self.params["probe_radius"])
            b_val = a_val * float(self.params["coax_ba_ratio"])
            wall_val = float(self.params["coax_wall_thickness"])
            Lp_val = float(self.params["coax_port_length"])
            h_sub_val = float(self.params["substrate_thickness"])
            clear_val = float(self.params["antipad_clearance"])
            if b_val - a_val < 0.02:
                b_val = a_val + 0.02

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
            if r_end <= r_start:
                r_end = a_val + 0.75 * (b_val - a_val)
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
    def _add_or_set_post_var(self, name: str, value: str) -> bool:
        """Tenta atualizar variável de pós-processamento; se não existir, cria."""
        try:
            # Primeiro tenta verificar se a variável já existe
            try:
                existing_vars = self.hfss.odesign.GetVariables()
                if name in existing_vars:
                    # Variável existe, vamos atualizá-la
                    self.hfss.odesign.ChangeProperty(
                        [
                            "NAME:AllTabs",
                            [
                                "NAME:LocalVariableTab",
                                ["NAME:PropServers", "LocalVariables"],
                                ["NAME:ChangedProps", ["NAME:" + name, "Value:=", value]]
                            ]
                        ]
                    )
                    self.log_message(f"Post var '{name}' updated to {value}.")
                    return True
            except:
                pass
            
            # Se não existe, cria uma nova
            self.hfss.odesign.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:LocalVariableTab",
                        ["NAME:PropServers", "LocalVariables"],
                        ["NAME:NewProps", ["NAME:" + name, "PropType:=", "VariableProp", "UserDef:=", True, "Value:=", value]]
                    ]
                ]
            )
            self.log_message(f"Post var '{name}' created = {value}.")
            return True
        except Exception as e:
            self.log_message(f"Add/Set post var '{name}' failed: {e}")
            return False

    def _edit_sources_with_vars(self, excitations: List[str], pvars: List[str], phvars: List[str]) -> bool:
        """Chama Solutions.EditSources ligando cada excitação a p_i (magnitude) e ph_i (fase)."""
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

    def _ensure_infinite_sphere(self, name="Infinite Sphere1") -> Optional[str]:
        """Cria (ou recria) Infinite Sphere com amostragem 1° x 1°."""
        try:
            rf = self.hfss.odesign.GetModule("RadField")
            # remove duplicata se já houver com mesmo nome
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
        """Cria p_i/ph_i e aplica em EditSources; popula UI de fontes."""
        try:
            exs = self._list_excitations()
            if not exs:
                self.log_message("No excitations found for post-processing.")
                return
                
            pvars, phvars = [], []
            for i in range(1, len(exs) + 1):
                p = f"p{i}"; ph = f"ph{i}"
                # Criar variáveis como variáveis de projeto (não de pós-processamento)
                self._add_or_set_post_var(p, "1W")
                self._add_or_set_post_var(ph, "0deg")
                pvars.append(p)
                phvars.append(ph)
                
            # Aplicar as fontes
            self._edit_sources_with_vars(exs, pvars, phvars)
            
            # Construir painel de fontes na UI
            self.populate_source_controls(exs)
        except Exception as e:
            self.log_message(f"Postprocess-after-solve error: {e}\n{traceback.format_exc()}")

    # ------------- Helpers de solução -------------
    def _fetch_solution(self, expression: str, setup_candidates: Optional[List[str]] = None, **kwargs):
        """Wrapper robusto para post.get_solution_data tentando diferentes nomes de setup/sweep."""
        if setup_candidates is None:
            setup_candidates = ["Setup1 : Sweep1", "Setup1:Sweep1", "Setup1 : LastAdaptive", "Setup1:LastAdaptive"]
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
        """Converte retorno de data_real() em ndarray 1D com tamanho esperado."""
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
        """Obtém nomes das excitações da simulação; ordena por índice Pn quando possível."""
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
        """Retorna (ângulo, ganho_dB) para corte Theta ou Phi em frequency GHz."""
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
        """Varre Phi (fixo) e pega Theta = All para montar grade 3D normalizada."""
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
                        # ignora phi