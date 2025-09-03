import os
import tempfile
import time
import ansys.aedt.core
import customtkinter as ctk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np
import math
import threading
import queue
from datetime import datetime
import json
import traceback
from typing import Tuple, List
import webbrowser

# ---------------- Aparência ----------------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ModernPatchAntennaDesigner:
    def __init__(self):
        self.hfss = None
        self.desktop = None
        self.temp_folder = None
        self.project_name = ""
        self.log_queue = queue.Queue()
        self.simulation_thread = None
        self.stop_simulation = False
        self.save_project = False
        self.is_simulation_running = False

        # ------- Parâmetros do usuário -------
        self.params = {
            "frequency": 10.0,             # GHz
            "gain": 12.0,                  # dBi
            "sweep_start": 8.0,            # GHz
            "sweep_stop": 12.0,            # GHz
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,
            "spacing_type": "lambda/2",
            "substrate_material": "Rogers RO4003C (tm)",
            "substrate_thickness": 0.5,    # mm
            "metal_thickness": 0.035,      # mm
            "er": 3.55,
            "tan_d": 0.0027,
            "feed_position": "edge",
            "probe_radius": 0.4,           # mm (a)
            "coax_er": 2.1,                # PTFE
            "coax_wall_thickness": 0.2,    # mm
            "coax_port_length": 3.0,       # mm  (Lp)
            "antipad_clearance": 0.0       # mm
        }

        # ------- Parâmetros calculados -------
        self.calculated_params = {
            "num_patches": 4,
            "spacing": 15.0,
            "patch_length": 9.57,
            "patch_width": 9.25,
            "rows": 2,
            "cols": 2,
            "lambda_g": 0.0,
            "feed_offset": 2.0,
            "substrate_width": 50.0,
            "substrate_length": 50.0
        }

        self.c = 3e8
        self.setup_gui()

    # ---------------- GUI ----------------
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        self.window.geometry("1600x1000")
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Configurar layout principal
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        # Criar menu superior
        self.setup_menu()

        # Header moderno
        header_frame = ctk.CTkFrame(self.window, height=100, fg_color=("#2B5B84", "#1E3A5F"), corner_radius=0)
        header_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        header_frame.grid_propagate(False)
        
        # Logo e título
        logo_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        logo_frame.pack(side="left", padx=20, pady=10)
        
        # Título
        title_label = ctk.CTkLabel(
            header_frame, 
            text="Patch Antenna Array Designer",
            font=ctk.CTkFont(size=28, weight="bold", family="Helvetica"),
            text_color="white"
        )
        title_label.pack(side="left", padx=20, pady=30)
        
        # Botões de ação no header
        action_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        action_frame.pack(side="right", padx=20, pady=10)
        
        ctk.CTkButton(action_frame, text="Help", command=self.show_help, 
                     width=80, height=30, fg_color="transparent", border_width=1,
                     text_color=("gray10", "gray90")).pack(side="left", padx=5)
        ctk.CTkButton(action_frame, text="About", command=self.show_about, 
                     width=80, height=30, fg_color="transparent", border_width=1,
                     text_color=("gray10", "gray90")).pack(side="left", padx=5)

        # Área principal com abas
        self.tabview = ctk.CTkTabview(self.window, fg_color=("gray92", "gray15"))
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        for tab_name in ["Design", "Simulation", "Results", "Log"]:
            self.tabview.add(tab_name)
            self.tabview.tab(tab_name).grid_columnconfigure(0, weight=1)

        self.setup_design_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()

        # Barra de status
        status_frame = ctk.CTkFrame(self.window, height=40, fg_color=("gray85", "gray20"))
        status_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 5))
        status_frame.grid_propagate(False)
        
        self.status_label = ctk.CTkLabel(
            status_frame, 
            text="Ready to calculate parameters",
            font=ctk.CTkFont(weight="bold")
        )
        self.status_label.pack(side="left", padx=15, pady=10)
        
        # Indicador de versão
        version_label = ctk.CTkLabel(
            status_frame, 
            text="v1.0 © 2025 Antenna Design Suite",
            font=ctk.CTkFont(size=10),
            text_color=("gray40", "gray60")
        )
        version_label.pack(side="right", padx=15, pady=10)

        self.process_log_queue()

    def setup_menu(self):
        # Criar menu tradicional (tkinter) para funcionalidades adicionais
        menu_bar = ctk.CTkFrame(self.window, height=30, fg_color=("gray90", "gray20"))
        menu_bar.grid(row=0, column=0, sticky="nwe", padx=10, pady=(5, 0))
        
        # Botões do menu
        menu_items = ["File", "Edit", "View", "Tools", "Help"]
        for i, item in enumerate(menu_items):
            btn = ctk.CTkButton(menu_bar, text=item, width=60, height=25, 
                               fg_color="transparent", hover_color=("gray80", "gray25"),
                               text_color=("gray20", "gray80"))
            btn.grid(row=0, column=i, padx=2, pady=2)
            
            # Adicionar funcionalidades básicas
            if item == "File":
                btn.configure(command=self.file_menu)
            elif item == "Help":
                btn.configure(command=self.show_help)

    def file_menu(self):
        # Menu de arquivo simples
        menu = ctk.CTkToplevel(self.window)
        menu.title("File Menu")
        menu.geometry("200x150")
        menu.transient(self.window)
        menu.grab_set()
        
        ctk.CTkButton(menu, text="New Project", command=lambda: self.log_message("New Project clicked", "INFO")).pack(pady=5)
        ctk.CTkButton(menu, text="Open Project", command=self.load_parameters).pack(pady=5)
        ctk.CTkButton(menu, text="Save Project", command=self.save_parameters).pack(pady=5)
        ctk.CTkButton(menu, text="Exit", command=self.on_closing).pack(pady=5)

    def create_section(self, parent, title, row, column, padx=10, pady=10, colspan=1):
        section_frame = ctk.CTkFrame(parent, fg_color=("gray96", "gray18"), corner_radius=8)
        section_frame.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady, columnspan=colspan)
        section_frame.grid_columnconfigure(0, weight=1)
        
        # Header da seção com ícone
        header_frame = ctk.CTkFrame(section_frame, fg_color=("gray90", "gray22"), corner_radius=6)
        header_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 0))
        
        title_label = ctk.CTkLabel(
            header_frame, 
            text=title,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("gray20", "gray80")
        )
        title_label.pack(side="left", padx=10, pady=5)
        
        return section_frame

    def setup_design_tab(self):
        tab = self.tabview.tab("Design")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame principal com scroll
        main_frame = ctk.CTkScrollableFrame(tab, fg_color=("gray92", "gray15"))
        main_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5, columnspan=2)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # Seção de parâmetros da antena (coluna 1)
        antenna_section = self.create_section(main_frame, "Antenna Parameters", 0, 0, pady=15)
        self.create_parameter_controls(antenna_section, [
            ("Central Frequency (GHz):", "frequency", self.params["frequency"]),
            ("Desired Gain (dBi):", "gain", self.params["gain"]),
            ("Sweep Start (GHz):", "sweep_start", self.params["sweep_start"]),
            ("Sweep Stop (GHz):", "sweep_stop", self.params["sweep_stop"]),
            ("Patch Spacing:", "spacing_type", self.params["spacing_type"], 
             ["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])
        ])

        # Seção de parâmetros do substrato (coluna 1)
        substrate_section = self.create_section(main_frame, "Substrate Parameters", 1, 0, pady=15)
        self.create_parameter_controls(substrate_section, [
            ("Substrate Material:", "substrate_material", self.params["substrate_material"],
             ["Rogers RO4003C (tm)", "FR4_epoxy", "Duroid (tm)", "Air"]),
            ("Relative Permittivity (εr):", "er", self.params["er"]),
            ("Loss Tangent (tan δ):", "tan_d", self.params["tan_d"]),
            ("Substrate Thickness (mm):", "substrate_thickness", self.params["substrate_thickness"]),
            ("Metal Thickness (mm):", "metal_thickness", self.params["metal_thickness"])
        ])

        # Seção de parâmetros do coaxial (coluna 2)
        coax_section = self.create_section(main_frame, "Coaxial Feed Parameters", 0, 1, pady=15)
        self.create_parameter_controls(coax_section, [
            ("Probe Radius a (mm):", "probe_radius", self.params["probe_radius"]),
            ("Coax εr (PTFE):", "coax_er", self.params["coax_er"]),
            ("Shield Wall (mm):", "coax_wall_thickness", self.params["coax_wall_thickness"]),
            ("Port Length below GND (mm):", "coax_port_length", self.params["coax_port_length"]),
            ("Anti-pad clearance (mm):", "antipad_clearance", self.params["antipad_clearance"]),
            ("Feed Position:", "feed_position", self.params["feed_position"], 
             ["edge", "inset"])
        ])

        # Seção de configurações de simulação (coluna 2)
        sim_section = self.create_section(main_frame, "Simulation Settings", 1, 1, pady=15)
        self.create_parameter_controls(sim_section, [
            ("CPU Cores:", "cores", self.params["cores"]),
            ("Show HFSS Interface:", "show_gui", not self.params["non_graphical"], None, True),
            ("Save Project:", "save_project", self.save_project, None, True)
        ])

        # Seção de parâmetros calculados (full width)
        calc_section = self.create_section(main_frame, "Calculated Parameters", 2, 0, pady=15, colspan=2)
        self.setup_calculated_params(calc_section)

        # Visualização do array (full width)
        viz_section = self.create_section(main_frame, "Array Visualization", 3, 0, pady=15, colspan=2)
        self.setup_array_visualization(viz_section)

    def create_parameter_controls(self, parent, parameters):
        self.entries = []
        grid_frame = ctk.CTkFrame(parent, fg_color="transparent")
        grid_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        grid_frame.grid_columnconfigure(1, weight=1)
        
        for i, (label, key, value, *extra) in enumerate(parameters):
            # Label
            ctk.CTkLabel(grid_frame, text=label, font=ctk.CTkFont(weight="bold")
                         ).grid(row=i, column=0, padx=5, pady=8, sticky="w")
            
            # Widget de entrada
            if extra and extra[0] is not None:
                # É um combobox
                if isinstance(extra[0], list):
                    var = ctk.StringVar(value=value)
                    widget = ctk.CTkComboBox(grid_frame, values=extra[0], variable=var, width=200)
                    widget.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                    self.entries.append((key, var))
                # É um checkbox
                elif extra[0] is True:
                    var = ctk.BooleanVar(value=value)
                    widget = ctk.CTkCheckBox(grid_frame, text="", variable=var, width=30)
                    widget.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    self.entries.append((key, var))
            else:
                # É um campo de entrada normal
                widget = ctk.CTkEntry(grid_frame, width=200)
                widget.insert(0, str(value))
                widget.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                self.entries.append((key, widget))

    def setup_calculated_params(self, parent):
        grid = ctk.CTkFrame(parent, fg_color="transparent")
        grid.grid(row=1, column=0, sticky="nsew", padx=15, pady=10)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(2, weight=1)

        # Títulos das colunas
        ctk.CTkLabel(grid, text="Parameter", font=ctk.CTkFont(weight="bold", size=12),
                     text_color=("gray30", "gray70")).grid(row=0, column=0, sticky="w", pady=5)
        ctk.CTkLabel(grid, text="Value", font=ctk.CTkFont(weight="bold", size=12),
                     text_color=("gray30", "gray70")).grid(row=0, column=1, sticky="w", pady=5)
        ctk.CTkLabel(grid, text="Unit", font=ctk.CTkFont(weight="bold", size=12),
                     text_color=("gray30", "gray70")).grid(row=0, column=2, sticky="w", pady=5)

        # Parâmetros calculados
        self.patches_label = self.create_param_row(grid, "Number of Patches", "4", "", 1)
        self.rows_cols_label = self.create_param_row(grid, "Configuration", "2 x 2", "", 2)
        self.spacing_label = self.create_param_row(grid, "Spacing", "15.0", "mm (lambda/2)", 3)
        self.dimensions_label = self.create_param_row(grid, "Patch Dimensions", "9.57 x 9.25", "mm", 4)
        self.lambda_label = self.create_param_row(grid, "Guided Wavelength", "0.0", "mm", 5)
        self.feed_offset_label = self.create_param_row(grid, "Feed Offset", "2.0", "mm", 6)
        self.substrate_dims_label = self.create_param_row(grid, "Substrate Dimensions", "0.00 x 0.00", "mm", 7)

        # Botões de ação
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=15)
        
        ctk.CTkButton(button_frame, text="Calculate Parameters", command=self.calculate_parameters,
                      fg_color="#2E8B57", hover_color="#3CB371", width=180, height=35,
                      font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Save Parameters", command=self.save_parameters,
                      fg_color="#4169E1", hover_color="#6495ED", width=140, height=35).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Load Parameters", command=self.load_parameters,
                      fg_color="#FF8C00", hover_color="#FFA500", width=140, height=35).pack(side="left", padx=10)

    def create_param_row(self, parent, name, value, unit, row):
        ctk.CTkLabel(parent, text=name, font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, sticky="w", pady=3)
        value_label = ctk.CTkLabel(parent, text=value, font=ctk.CTkFont(weight="bold"))
        value_label.grid(row=row, column=1, sticky="w", pady=3)
        ctk.CTkLabel(parent, text=unit).grid(row=row, column=2, sticky="w", pady=3)
        return value_label

    def setup_array_visualization(self, parent):
        # Frame para visualização do array
        viz_frame = ctk.CTkFrame(parent, fg_color=("gray90", "gray20"), height=200)
        viz_frame.grid(row=1, column=0, sticky="nsew", padx=15, pady=15)
        viz_frame.grid_propagate(False)
        
        # Placeholder para visualização
        placeholder = ctk.CTkLabel(viz_frame, text="Array visualization will be shown here\nafter parameters calculation",
                                  font=ctk.CTkFont(size=12, slant="italic"),
                                  text_color=("gray40", "gray60"))
        placeholder.pack(expand=True)
        
        # Botão para atualizar visualização
        ctk.CTkButton(parent, text="Update Visualization", command=self.update_visualization,
                     width=160, height=30).grid(row=2, column=0, pady=(0, 15))

    def update_visualization(self):
        # Placeholder para futura implementação de visualização do array
        self.log_message("Visualization update requested", "INFO")

    def setup_simulation_tab(self):
        tab = self.tabview.tab("Simulation")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(tab, fg_color=("gray92", "gray15"))
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)

        # Título
        title_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray20"), corner_radius=8)
        title_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=10)
        
        ctk.CTkLabel(title_frame, text="Simulation Control", 
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=12)

        # Controles de simulação
        control_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        control_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        control_frame.grid_columnconfigure(0, weight=1)
        control_frame.grid_columnconfigure(1, weight=1)

        # Botões de controle
        btn_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=0, columnspan=2, pady=20)
        
        self.run_button = ctk.CTkButton(btn_frame, text="▶ Start Simulation",
                                        command=self.start_simulation_thread,
                                        fg_color="#2E8B57", hover_color="#3CB371",
                                        height=45, width=180, font=ctk.CTkFont(size=14, weight="bold"))
        self.run_button.pack(side="left", padx=10)

        self.stop_button = ctk.CTkButton(btn_frame, text="⏹ Stop Simulation",
                                         command=self.stop_simulation_thread,
                                         fg_color="#DC143C", hover_color="#FF4500",
                                         height=45, width=180, state="disabled",
                                         font=ctk.CTkFont(size=14, weight="bold"))
        self.stop_button.pack(side="left", padx=10)

        # Barra de progresso
        progress_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        progress_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)

        ctk.CTkLabel(progress_frame, text="Simulation Progress:",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w")

        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=20, progress_color="#4B9CD3")
        self.progress_bar.pack(fill="x", pady=5)
        self.progress_bar.set(0)

        # Status da simulação
        status_frame = ctk.CTkFrame(control_frame, fg_color=("gray90", "gray20"), corner_radius=8)
        status_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)
        
        self.sim_status_label = ctk.CTkLabel(status_frame, text="Simulation not started",
                                             font=ctk.CTkFont(weight="bold"))
        self.sim_status_label.pack(pady=12)

        # Informações adicionais
        info_frame = ctk.CTkFrame(control_frame, fg_color=("gray94", "gray18"), corner_radius=8)
        info_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)

        ctk.CTkLabel(info_frame,
                     text="ℹ️ Note: Simulation may take several minutes depending on array size and computer resources",
                     font=ctk.CTkFont(size=12),
                     text_color=("gray40", "gray60")).pack(padx=15, pady=12)

        # Estatísticas de tempo (placeholder)
        stats_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        stats_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=5)
        
        ctk.CTkLabel(stats_frame, text="Estimated time: 15-30 min", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).pack(side="left")
        
        ctk.CTkLabel(stats_frame, text="Memory required: ~4 GB", font=ctk.CTkFont(size=11),
                     text_color=("gray40", "gray60")).pack(side="right")

    def setup_results_tab(self):
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(tab, fg_color=("gray92", "gray15"))
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        # Cabeçalho
        header_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray20"), corner_radius=8)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        ctk.CTkLabel(header_frame, text="Simulation Results",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=12)

        # Área de gráficos
        graph_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray18"))
        graph_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)

        # Configurar figura do matplotlib
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.fig.set_facecolor('#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF')
        self.ax.set_facecolor('#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF')

        if ctk.get_appearance_mode() == "Dark":
            self.ax.tick_params(colors='white')
            self.ax.xaxis.label.set_color('white')
            self.ax.yaxis.label.set_color('white')
            self.ax.title.set_color('white')
            for side in ['bottom','top','right','left']:
                self.ax.spines[side].set_color('white')
            self.ax.grid(color='gray')

        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)

        # Controles de exportação
        export_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        export_frame.grid(row=2, column=0, pady=10)
        
        ctk.CTkButton(export_frame, text="📊 Export CSV", command=self.export_csv,
                      fg_color="#6A5ACD", hover_color="#7B68EE", width=120).pack(side="left", padx=10)
        ctk.CTkButton(export_frame, text="🖼️ Export PNG", command=self.export_png,
                      fg_color="#20B2AA", hover_color="#40E0D0", width=120).pack(side="left", padx=10)
        ctk.CTkButton(export_frame, text="📄 Export Report", command=self.export_report,
                      fg_color="#FF6347", hover_color="#FF7F50", width=120).pack(side="left", padx=10)

    def setup_log_tab(self):
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(tab, fg_color=("gray92", "gray15"))
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        # Cabeçalho
        header_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray20"), corner_radius=8)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        ctk.CTkLabel(header_frame, text="Simulation Log",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=12)

        # Área de log
        log_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray18"))
        log_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(log_frame, width=900, height=500, 
                                      font=ctk.CTkFont(family="Consolas", size=12),
                                      fg_color=("gray95", "gray12"))
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n" + "="*50 + "\n")

        # Controles de log
        control_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        control_frame.grid(row=2, column=0, pady=10)
        
        ctk.CTkButton(control_frame, text="🗑️ Clear Log", command=self.clear_log,
                      fg_color="#696969", hover_color="#808080").pack(side="left", padx=10)
        ctk.CTkButton(control_frame, text="💾 Save Log", command=self.save_log,
                      fg_color="#4169E1", hover_color="#6495ED").pack(side="left", padx=10)
        ctk.CTkButton(control_frame, text="🔍 Search", command=self.search_log,
                      fg_color="#228B22", hover_color="#32CD32").pack(side="left", padx=10)

    # ------------- utilitários de log -------------
    def log_message(self, message, level="INFO"):
        timestamp = datetime.now().strftime('%H:%M:%S')
        levels = {
            "INFO": ("", "black"),
            "WARNING": ("⚠️ ", "orange"),
            "ERROR": ("❌ ", "red"),
            "SUCCESS": ("✅ ", "green")
        }
        
        emoji, color = levels.get(level, ("", "black"))
        formatted_msg = f"[{timestamp}] {emoji}{message}\n"
        self.log_queue.put((formatted_msg, color))

    def process_log_queue(self):
        try:
            while True:
                msg, color = self.log_queue.get_nowait()
                self.log_text.insert("end", msg)
                self.log_text.see("end")
        except queue.Empty:
            pass
        finally:
            if self.window.winfo_exists():
                self.window.after(100, self.process_log_queue)

    def clear_log(self):
        self.log_text.delete("1.0", "end")
        self.log_message("Log cleared", "INFO")

    def save_log(self):
        try:
            with open("simulation_log.txt", "w", encoding="utf-8") as f:
                f.write(self.log_text.get("1.0", "end"))
            self.log_message("Log saved to simulation_log.txt", "SUCCESS")
        except Exception as e:
            self.log_message(f"Error saving log: {str(e)}", "ERROR")

    def search_log(self):
        # Placeholder para funcionalidade de busca
        self.log_message("Search functionality not yet implemented", "WARNING")

    def export_csv(self):
        try:
            if hasattr(self, 'simulation_data'):
                np.savetxt("simulation_results.csv", self.simulation_data, delimiter=",",
                           header="Frequency (GHz), S11 (dB)", comments='')
                self.log_message("Data exported to simulation_results.csv", "SUCCESS")
            else:
                self.log_message("No simulation data available for export", "WARNING")
        except Exception as e:
            self.log_message(f"Error exporting CSV: {str(e)}", "ERROR")

    def export_png(self):
        try:
            if hasattr(self, 'fig'):
                self.fig.savefig("simulation_results.png", dpi=300, bbox_inches='tight',
                                facecolor=self.fig.get_facecolor())
                self.log_message("Plot saved to simulation_results.png", "SUCCESS")
        except Exception as e:
            self.log_message(f"Error saving plot: {str(e)}", "ERROR")

    def export_report(self):
        # Placeholder para exportação de relatório completo
        self.log_message("Report export functionality not yet implemented", "WARNING")

    def show_help(self):
        webbrowser.open("https://github.com/ansys/pyaedt")
        
    def show_about(self):
        about_window = ctk.CTkToplevel(self.window)
        about_window.title("About")
        about_window.geometry("400x300")
        about_window.transient(self.window)
        about_window.grab_set()
        
        ctk.CTkLabel(about_window, text="Patch Antenna Array Designer", 
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        
        ctk.CTkLabel(about_window, text="Version 1.0", 
                     font=ctk.CTkFont(size=14)).pack(pady=5)
        
        ctk.CTkLabel(about_window, text="A professional tool for designing and simulating\npatch antenna arrays using PyAEDT and HFSS",
                     font=ctk.CTkFont(size=12), justify="center").pack(pady=10)
        
        ctk.CTkLabel(about_window, text="© 2025 Antenna Design Suite\nAll rights reserved",
                     font=ctk.CTkFont(size=10), text_color=("gray40", "gray60")).pack(pady=20)
        
        ctk.CTkButton(about_window, text="Close", command=about_window.destroy).pack(pady=10)

    # ----------- Física / cálculos -----------
    def get_parameters(self):
        self.log_message("Getting parameters from interface", "INFO")
        for key, widget in self.entries:
            try:
                if key == "cores":
                    self.params[key] = int(widget.get()) if isinstance(widget, ctk.CTkEntry) else int(self.params[key])
                elif key == "show_gui":
                    self.params["non_graphical"] = not widget.get()
                elif key == "save_project":
                    self.save_project = widget.get()
                elif key in ["substrate_thickness", "metal_thickness", "er", "tan_d",
                             "probe_radius", "coax_er", "coax_wall_thickness",
                             "coax_port_length", "antipad_clearance"]:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
                elif key in ["spacing_type", "substrate_material", "feed_position"]:
                    self.params[key] = widget.get()
                else:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
            except Exception as e:
                msg = f"Invalid value for {key}: {str(e)}"
                self.status_label.configure(text=msg)
                self.log_message(msg, "ERROR")
                return False
        self.log_message("All parameters retrieved successfully", "SUCCESS")
        return True

    def calculate_patch_dimensions(self, frequency_ghz: float) -> Tuple[float, float, float]:
        f = frequency_ghz * 1e9
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"]) / 1000.0
        W = self.c / (2 * f) * math.sqrt(2 / (er + 1))
        eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W) ** (-0.5)
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        L_eff = self.c / (2 * f * math.sqrt(eeff))
        L = L_eff - 2 * dL
        lambda_g_mm = (self.c / (f * math.sqrt(eeff))) * 1000.0
        return (L * 1000.0, W * 1000.0, lambda_g_mm)

    def calculate_substrate_size(self):
        L = self.calculated_params["patch_length"]
        W = self.calculated_params["patch_width"]
        s = self.calculated_params["spacing"]
        r = self.calculated_params["rows"]
        c = self.calculated_params["cols"]
        total_w = c * W + (c - 1) * s
        total_l = r * L + (r - 1) * s
        margin = max(total_w, total_l) * 0.2
        self.calculated_params["substrate_width"] = total_w + 2 * margin
        self.calculated_params["substrate_length"] = total_l + 2 * margin
        self.log_message(f"Substrate size calculated: {self.calculated_params['substrate_width']:.2f} x {self.calculated_params['substrate_length']:.2f} mm", "INFO")

    def calculate_parameters(self):
        self.log_message("Starting parameter calculation", "INFO")
        if not self.get_parameters():
            self.log_message("Parameter calculation failed due to invalid input", "ERROR")
            return
        try:
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params.update({"patch_length": L_mm, "patch_width": W_mm, "lambda_g": lambda_g_mm})
            self.calculated_params["feed_offset"] = 0.1 * W_mm

            # --- Número de elementos pela meta de ganho ---
            G_elem = 8.0
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
            self.calculated_params.update({"num_patches": rows * cols, "rows": rows, "cols": cols})
            self.log_message(f"Array sizing -> target gain {G_des} dBi, N_req≈{N_req}, layout {rows}x{cols} (= {rows*cols} patches)", "INFO")

            # --- Espaçamento ---
            lambda0_m = self.c / (self.params["frequency"] * 1e9)
            factors = {"lambda/2": 0.5, "lambda": 1.0, "0.7*lambda": 0.7, "0.8*lambda": 0.8, "0.9*lambda": 0.9}
            spacing_mm = factors.get(self.params["spacing_type"], 0.5) * lambda0_m * 1000.0
            self.calculated_params["spacing"] = spacing_mm

            self.calculate_substrate_size()

            self.patches_label.configure(text=f"Number of Patches: {rows*cols}")
            self.rows_cols_label.configure(text=f"Configuration: {rows} x {cols}")
            self.spacing_label.configure(text=f"Spacing: {spacing_mm:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch Dimensions: {L_mm:.2f} x {W_mm:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {lambda_g_mm:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset: {self.calculated_params['feed_offset']:.2f} mm")
            self.substrate_dims_label.configure(text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x {self.calculated_params['substrate_length']:.2f} mm")
            self.status_label.configure(text="Parameters calculated successfully!")
            self.log_message("Parameters calculated successfully", "SUCCESS")
        except Exception as e:
            msg = f"Error in calculation: {str(e)}"
            self.status_label.configure(text=msg)
            self.log_message(msg, "ERROR")
            self.log_message(f"Traceback: {traceback.format_exc()}", "ERROR")

    # --------- utilidades de modelagem ---------
    def _ensure_material(self, name: str, er: float, tan_d: float):
        try:
            if not self.hfss.materials.checkifmaterialexists(name):
                self.hfss.materials.add_material(name)
                m = self.hfss.materials.material_keys[name]
                m.permittivity = er
                m.dielectric_loss_tangent = tan_d
                self.log_message(f"Created material: {name} (er={er}, tanδ={tan_d})", "INFO")
        except Exception as e:
            self.log_message(f"Material management warning for '{name}': {e}", "WARNING")

    def _set_design_variables(self, L, W, spacing, rows, cols, h_sub, sub_w, sub_l):
        a = float(self.params["probe_radius"])
        er_cx = float(self.params["coax_er"])
        wall = float(self.params["coax_wall_thickness"])
        Lp = float(self.params["coax_port_length"])
        clear = float(self.params["antipad_clearance"])
        b = a * math.exp(50.0 * math.sqrt(er_cx) / 60.0)  # 50Ω coax formula

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
        self.hfss["probeK"] = "0.3"
        self.hfss["padAir"] = f"{max(spacing, W, L)/2 + Lp + 2.0}mm"

        return a, b, wall, Lp, clear

    def _create_coax_feed_lumped(self, ground, substrate, x_feed: float, y_feed: float,
                                 name_prefix: str):
        """
        Constrói pino, PTFE anular, blindagem, furo no substrato e anti-pad.
        Cria Lumped Port em z = -Lp como anel (raios a..b) com linha de integração radial.
        """
        try:
            # ---- parâmetros numéricos ----
            a_val = float(self.params["probe_radius"])
            Lp_val = float(self.params["coax_port_length"])
            h_sub_val = float(self.params["substrate_thickness"])
            b_val = a_val * math.exp(50.0 * math.sqrt(float(self.params["coax_er"])) / 60.0)
            wall_val = float(self.params["coax_wall_thickness"])
            clear_val = float(self.params["antipad_clearance"])

            if (b_val - a_val) < 0.02:
                b_val = a_val + 0.02

            # ---- PINO ----
            pin = self.hfss.modeler.create_cylinder(
                orientation="Z",
                origin=[x_feed, y_feed, -Lp_val],
                radius=a_val,
                height=h_sub_val + Lp_val + 0.001,
                name=f"{name_prefix}_Pin",
                material="copper"
            )

            # ---- PTFE (anel) em -Lp..0 (mantendo o pino!) ----
            ptfe_solid = self.hfss.modeler.create_cylinder(
                orientation="Z",
                origin=[x_feed, y_feed, -Lp_val],
                radius=b_val,
                height=Lp_val,
                name=f"{name_prefix}_PTFEsolid",
                material="PTFE_Custom"
            )
            self.hfss.modeler.subtract(ptfe_solid, [pin], keep_originals=True)
            ptfe = ptfe_solid
            ptfe.name = f"{name_prefix}_PTFE"

            # ---- BLINDAGEM (tubo) ----
            shield_outer = self.hfss.modeler.create_cylinder(
                orientation="Z",
                origin=[x_feed, y_feed, -Lp_val],
                radius=b_val + wall_val,
                height=Lp_val,
                name=f"{name_prefix}_ShieldOuter",
                material="copper"
            )
            shield_inner_void = self.hfss.modeler.create_cylinder(
                orientation="Z",
                origin=[x_feed, y_feed, -Lp_val],
                radius=b_val,
                height=Lp_val,
                name=f"{name_prefix}_ShieldInnerVoid",
                material="vacuum"
            )
            self.hfss.modeler.subtract(shield_outer, [shield_inner_void], keep_originals=False)
            shield = shield_outer

            # ---- FURO no substrato + anti-pad no GND ----
            hole_r = b_val + clear_val
            sub_hole = self.hfss.modeler.create_cylinder(
                orientation="Z",
                origin=[x_feed, y_feed, 0.0],
                radius=hole_r,
                height=h_sub_val,
                name=f"{name_prefix}_SubHole",
                material="vacuum"
            )
            self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=False)
            g_hole = self.hfss.modeler.create_circle(
                orientation="XY",
                origin=[x_feed, y_feed, 0.0],
                radius=hole_r,
                name=f"{name_prefix}_GndHole",
                material="vacuum"
            )
            self.hfss.modeler.subtract(ground, [g_hole], keep_originals=False)

            # ---- SHEET do porto (anel entre a e b) no plano z=-Lp ----
            port_ring = self.hfss.modeler.create_circle(
                orientation="XY",
                origin=[x_feed, y_feed, -Lp_val],
                radius=b_val,
                name=f"{name_prefix}_PortRing",
                material="vacuum"
            )
            port_hole = self.hfss.modeler.create_circle(
                orientation="XY",
                origin=[x_feed, y_feed, -Lp_val],
                radius=a_val,
                name=f"{name_prefix}_PortHole",
                material="vacuum"
            )
            self.hfss.modeler.subtract(port_ring, [port_hole], keep_originals=False)

            # ---- Linha de integração: dois pontos DENTRO do anel ----
            eps_line = min(0.1 * (b_val - a_val), 0.05)  # mm
            r_start = a_val + eps_line
            r_end = b_val - eps_line
            if r_end <= r_start:
                r_end = a_val + 0.75 * (b_val - a_val)

            int_start = [x_feed + r_start, y_feed, -Lp_val]
            int_end = [x_feed + r_end,   y_feed, -Lp_val]

            port = self.hfss.lumped_port(
                assignment=port_ring.name,
                integration_line=[int_start, int_end],
                impedance=50.0,
                name=f"{name_prefix}_Lumped",
                renormalize=True
            )
            self.log_message(f"Lumped Port '{name_prefix}_Lumped' created (integration line).", "INFO")
            return pin, ptfe, shield

        except Exception as e:
            self.log_message(f"Exception in coax creation '{name_prefix}': {e}", "ERROR")
            self.log_message(f"Traceback: {traceback.format_exc()}", "ERROR")
            return None, None, None

    # ------------- Simulação -------------
    def start_simulation_thread(self):
        if self.is_simulation_running:
            self.log_message("Simulation is already running", "WARNING")
            return
        self.stop_simulation = False
        self.is_simulation_running = True
        threading.Thread(target=self.run_simulation, daemon=True).start()

    def stop_simulation_thread(self):
        self.stop_simulation = True
        self.log_message("Simulation stop requested", "INFO")

    def run_simulation(self):
        try:
            self.log_message("Starting simulation", "INFO")
            self.run_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            self.sim_status_label.configure(text="Simulation in progress")
            self.progress_bar.set(0)

            if not self.get_parameters():
                self.log_message("Invalid parameters. Aborting.", "ERROR")
                return
            if self.calculated_params["num_patches"] < 1:
                self.calculate_parameters()

            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
            self.project_name = os.path.join(self.temp_folder.name, "patch_array.aedt")
            self.log_message(f"Creating project: {self.project_name}", "INFO")
            self.progress_bar.set(0.1)

            self.log_message("Initializing HFSS", "INFO")
            self.desktop = ansys.aedt.core.Desktop(
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"],
                new_desktop=True
            )
            self.progress_bar.set(0.2)

            self.log_message("Creating HFSS project", "INFO")
            self.hfss = ansys.aedt.core.Hfss(
                project=self.project_name,
                design="patch_array",
                solution_type="DrivenModal",
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"]
            )
            self.log_message("HFSS initialized successfully", "SUCCESS")
            self.progress_bar.set(0.3)

            self.hfss.modeler.model_units = "mm"
            self.log_message("Model units set to: mm", "INFO")

            # Materiais
            sub_mat = self.params["substrate_material"]
            er = float(self.params["er"])
            tan_d = float(self.params["tan_d"])
            if sub_mat not in ["Rogers RO4003C (tm)", "FR4_epoxy", "Duroid (tm)", "Air"]:
                sub_mat = "Custom_Substrate"
            self._ensure_material(sub_mat, er, tan_d)
            self._ensure_material("PTFE_Custom", float(self.params["coax_er"]), 0.0002)

            # Geometria / variáveis
            L = float(self.calculated_params["patch_length"])
            W = float(self.calculated_params["patch_width"])
            spacing = float(self.calculated_params["spacing"])
            rows = int(self.calculated_params["rows"])
            cols = int(self.calculated_params["cols"])
            h_sub = float(self.params["substrate_thickness"])
            sub_w = float(self.calculated_params["substrate_width"])
            sub_l = float(self.calculated_params["substrate_length"])

            a, b, wall, Lp, clear = self._set_design_variables(L, W, spacing, rows, cols, h_sub, sub_w, sub_l)

            # Substrato e Ground
            self.log_message("Creating substrate", "INFO")
            substrate = self.hfss.modeler.create_box(
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL", "h_sub"],
                name="Substrate",
                material=sub_mat
            )
            self.log_message("Creating ground plane", "INFO")
            ground = self.hfss.modeler.create_rectangle(
                orientation="XY",
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL"],
                name="Ground",
                material="copper"
            )

            # Patches
            self.log_message(f"Creating {rows*cols} patches in {rows}x{cols} configuration", "INFO")
            patches: List = []
            total_width = cols * W + (cols - 1) * spacing
            total_length = rows * L + (rows - 1) * spacing
            start_x = -total_width / 2 + W / 2
            start_y = -total_length / 2 + L / 2

            self.progress_bar.set(0.4)
            count = 0
            for r in range(rows):
                for c in range(cols):
                    if self.stop_simulation:
                        self.log_message("Simulation stopped by user", "INFO")
                        return
                    count += 1
                    patch_name = f"Patch_{count}"
                    cx = start_x + c * (W + spacing)
                    cy = start_y + r * (L + spacing)

                    origin = [cx - W / 2, cy - L / 2, "h_sub"]
                    self.log_message(f"Creating patch {count} at ({r}, {c})", "INFO")

                    patch = self.hfss.modeler.create_rectangle(
                        orientation="XY",
                        origin=origin,
                        sizes=["patchW", "patchL"],
                        name=patch_name,
                        material="copper"
                    )
                    patches.append(patch)

                    # ---- Pad e coax com coordenadas NUMÉRICAS ----
                    y_feed = cy - 0.5*L + 0.3*L       # 30% de inset
                    pad = self.hfss.modeler.create_circle(
                        orientation="XY",
                        origin=[cx, y_feed, "h_sub"],
                        radius="a",
                        name=f"{patch_name}_Pad",
                        material="copper"
                    )
                    try:
                        self.hfss.modeler.unite([patch, pad])
                    except Exception:
                        pass

                    # Coax + Lumped Port
                    x_feed = cx
                    pin, ptfe, shield = self._create_coax_feed_lumped(
                        ground=ground,
                        substrate=substrate,
                        x_feed=x_feed,
                        y_feed=y_feed,
                        name_prefix=f"P{count}"
                    )

                    self.progress_bar.set(0.4 + 0.2 * (count / float(rows * cols)))

            if self.stop_simulation:
                self.log_message("Simulation stopped by user", "INFO")
                return

            # -------- Perfect E nos condutores de folha --------
            try:
                sheet_names = [ground.name] + [p.name for p in patches]
                self.hfss.assign_perfecte_to_sheets(sheet_names)
                self.log_message(f"Assigned PerfectE to sheets: {sheet_names}", "INFO")
            except Exception as e:
                self.log_message(f"PerfectE assignment warning: {e}", "WARNING")

            # Região de ar + radiação (λ/4)
            self.log_message("Creating air region + radiation boundary", "INFO")
            lambda0 = self.c / (self.params["sweep_start"] * 1e9) * 1000.0  # mm (NUMÉRICO)
            region_size = float(lambda0) / 4.0
            region = self.hfss.modeler.create_region(
                [region_size, region_size, region_size, region_size, region_size, region_size],
                is_percentage=False
            )
            self.hfss.assign_radiation_boundary_to_objects(region)
            self.progress_bar.set(0.7)

            # Setup + Sweep
            self.log_message("Creating simulation setup", "INFO")
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02

            self.log_message("Creating frequency sweep (linear step for 201 points)", "INFO")
            step = (self.params["sweep_stop"] - self.params["sweep_start"]) / 200.0
            try:
                setup.create_linear_step_sweep(
                    unit="GHz",
                    start_frequency=self.params["sweep_start"],
                    stop_frequency=self.params["sweep_stop"],
                    step_size=step,
                    name="Sweep1"
                )
            except Exception as e:
                self.log_message(f"Linear-step helper not available ({e}). Using interpolating sweep.", "WARNING")
                setup.create_frequency_sweep(
                    unit="GHz",
                    name="Sweep1",
                    start_frequency=self.params["sweep_start"],
                    stop_frequency=self.params["sweep_stop"],
                    sweep_type="Interpolating"
                )

            # Malha leve nos patches
            self.log_message("Assigning local mesh refinement", "INFO")
            try:
                lambda_g_mm = max(1e-6, self.calculated_params["lambda_g"])
                edge_len = max(lambda_g_mm / 60.0, W / 200.0)
                for p in patches:
                    self.hfss.mesh.assign_length_mesh([p], maximum_length=f"{edge_len}mm")
            except Exception as e:
                self.log_message(f"Mesh refinement warning: {e}", "WARNING")

            # --- Verificação de excitações ---
            try:
                exs = self.hfss.get_excitations_name() or []
            except Exception:
                exs = list(getattr(self.hfss, "excitations", []) or [])
            self.log_message(f"Excitations created: {len(exs)} -> {exs}", "INFO")
            if not exs:
                self.sim_status_label.configure(text="No excitations defined")
                self.log_message("No excitations found. Aborting before solve.", "ERROR")
                return

            self.log_message("Validating design", "INFO")
            _ = self.hfss.validate_full_design()

            self.log_message("Starting analysis", "INFO")
            self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])

            if self.stop_simulation:
                self.log_message("Simulation stopped by user", "INFO")
                return

            self.progress_bar.set(0.9)
            self.log_message("Processing results", "INFO")
            self.plot_results()
            self.progress_bar.set(1.0)
            self.sim_status_label.configure(text="Simulation completed")
            self.log_message("Simulation completed successfully", "SUCCESS")

        except Exception as e:
            msg = f"Error in simulation: {str(e)}"
            self.log_message(msg, "ERROR")
            self.sim_status_label.configure(text=msg)
            self.log_message(f"Traceback: {traceback.format_exc()}", "ERROR")
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    def plot_results(self):
        try:
            self.log_message("Plotting results", "INFO")
            self.ax.clear()

            # construir expressão baseada no nome do primeiro terminal
            try:
                exs = self.hfss.get_excitations_name() or []
            except Exception:
                exs = []
            expr = "dB(S(1,1))"
            if exs:
                p = exs[0].split(":")[0]
                expr = f"dB(S({p},{p}))"

            report = self.hfss.post.reports_by_category.standard(expressions=[expr])
            report.context = ["Setup1: Sweep1"]
            sol = report.get_solution_data()

            if sol:
                freqs = np.array(sol.primary_sweep_values, dtype=float)
                data = sol.data_real()
                if len(data) > 0:
                    s11 = np.array(data[0], dtype=float)
                    self.simulation_data = np.column_stack((freqs, s11))
                    self.ax.plot(freqs, s11, label=expr, linewidth=2)
                    self.ax.axhline(y=-10, linestyle='--', alpha=0.7, label='-10 dB')
                    self.ax.set_xlabel("Frequency (GHz)")
                    self.ax.set_ylabel("S-Parameter (dB)")
                    self.ax.set_title("S11 - Coax-fed Patch Array (Lumped Ports)")
                    self.ax.legend()
                    self.ax.grid(True)
                    cf = float(self.params["frequency"])
                    self.ax.axvline(x=cf, linestyle='--', alpha=0.7)
                    self.ax.text(cf + 0.1, self.ax.get_ylim()[1] - 2, f'{cf} GHz')
                    self.canvas.draw()
                    self.log_message("Results plotted successfully", "SUCCESS")
                    return

            # fallback
            self.log_message("No data from named-port expression, trying S(1,1)", "WARNING")
            report = self.hfss.post.reports_by_category.standard(expressions=["dB(S(1,1))"])
            report.context = ["Setup1: Sweep1"]
            sol = report.get_solution_data()
            if sol:
                freqs = np.array(sol.primary_sweep_values, dtype=float)
                data = sol.data_real()
                if len(data) > 0:
                    s11 = np.array(data[0], dtype=float)
                    self.simulation_data = np.column_stack((freqs, s11))
                    self.ax.plot(freqs, s11, label="dB(S(1,1))", linewidth=2)
                    self.ax.axhline(y=-10, linestyle='--', alpha=0.7, label='-10 dB')
                    self.ax.set_xlabel("Frequency (GHz)")
                    self.ax.set_ylabel("S-Parameter (dB)")
                    self.ax.set_title("S11 - Coax-fed Patch Array (Lumped Ports)")
                    self.ax.legend()
                    self.ax.grid(True)
                    self.canvas.draw()
                    self.log_message("Results plotted successfully (fallback)", "SUCCESS")
                else:
                    self.log_message("No S11 data available for plotting", "WARNING")
            else:
                self.log_message("Could not get simulation data", "ERROR")
        except Exception as e:
            self.log_message(f"Error plotting results: {str(e)}", "ERROR")
            self.log_message(f"Traceback: {traceback.format_exc()}", "ERROR")

    # ------------- Encerramento -------------
    def cleanup(self):
        try:
            if self.hfss and hasattr(self.hfss, 'close_project'):
                try:
                    if self.save_project:
                        self.hfss.save_project()
                    else:
                        self.hfss.close_project(save=False)
                except Exception as e:
                    self.log_message(f"Error closing project: {str(e)}", "ERROR")
            if self.desktop and hasattr(self.desktop, 'release_desktop'):
                try:
                    self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception as e:
                    self.log_message(f"Error releasing desktop: {str(e)}", "ERROR")
            if self.temp_folder and not self.save_project:
                try:
                    self.temp_folder.cleanup()
                except Exception as e:
                    self.log_message(f"Error cleaning up temporary files: {str(e)}", "ERROR")
        except Exception as e:
            self.log_message(f"Error during cleanup: {str(e)}", "ERROR")

    def on_closing(self):
        self.log_message("Application closing...", "INFO")
        self.cleanup()
        self.window.quit()
        self.window.destroy()

    def save_parameters(self):
        try:
            all_params = {**self.params, **self.calculated_params}
            with open("antenna_parameters.json", "w") as f:
                json.dump(all_params, f, indent=4)
            self.log_message("Parameters saved to antenna_parameters.json", "SUCCESS")
        except Exception as e:
            self.log_message(f"Error saving parameters: {str(e)}", "ERROR")

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
            self.log_message("Parameters loaded from antenna_parameters.json", "SUCCESS")
        except Exception as e:
            self.log_message(f"Error loading parameters: {str(e)}", "ERROR")

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
            self.rows_cols_label.configure(text=f"Configuration: {self.calculated_params['rows']} x {self.calculated_params['cols']}")
            self.spacing_label.configure(text=f"Spacing: {self.calculated_params['spacing']:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch Dimensions: {self.calculated_params['patch_length']:.2f} x {self.calculated_params['patch_width']:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {self.calculated_params['lambda_g']:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset: {self.calculated_params['feed_offset']:.2f} mm")
            self.substrate_dims_label.configure(text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x {self.calculated_params['substrate_length']:.2f} mm")
            self.log_message("Interface updated with loaded parameters", "INFO")
        except Exception as e:
            self.log_message(f"Error updating interface: {str(e)}", "ERROR")

    def run(self):
        try:
            self.window.mainloop()
        except Exception as e:
            self.log_message(f"Unexpected error: {str(e)}", "ERROR")
        finally:
            self.cleanup()

if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()