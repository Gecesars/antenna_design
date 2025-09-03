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

# Configuração do tema
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
            "frequency": 10.0,
            "gain": 12.0,
            "sweep_start": 8.0,
            "sweep_stop": 12.0,
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,
            "spacing_type": "lambda/2",
            "substrate_material": "Rogers RO4003C (tm)",
            "substrate_thickness": 0.5,   # mm
            "metal_thickness": 0.035,     # mm
            "er": 3.55,
            "tan_d": 0.0027,
            "feed_position": "edge",
            "probe_radius": 0.4,            # a (mm) - pino interno
            "coax_er": 2.1,                 # PTFE
            "coax_wall_thickness": 0.2,     # esp. blindagem (mm)
            "coax_port_length": 3.0,        # comprimento sob o GND (mm)
            "antipad_clearance": 0.0        # folga extra no furo do substrato (mm)
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
        self.window.geometry("1400x950")
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Configurar layout principal
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)
        
        # Header moderno
        header_frame = ctk.CTkFrame(self.window, height=80, fg_color=("gray85", "gray20"))
        header_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header_frame.grid_propagate(False)
        
        title_label = ctk.CTkLabel(
            header_frame, 
            text="Patch Antenna Array Designer",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=("gray10", "gray90")
        )
        title_label.pack(pady=20)
        
        # Tabview principal
        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Adicionar abas
        tabs = ["Design Parameters", "Simulation", "Results", "Log"]
        for tab_name in tabs:
            self.tabview.add(tab_name)
            self.tabview.tab(tab_name).grid_columnconfigure(0, weight=1)
        
        # Configurar cada aba
        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()
        
        # Barra de status
        status_frame = ctk.CTkFrame(self.window, height=40)
        status_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 5))
        status_frame.grid_propagate(False)
        
        self.status_label = ctk.CTkLabel(
            status_frame, 
            text="Ready to calculate parameters",
            font=ctk.CTkFont(weight="bold")
        )
        self.status_label.pack(pady=10)
        
        # Iniciar processamento de log
        self.process_log_queue()

    def create_section(self, parent, title, row, column, padx=10, pady=10):
        """Cria uma seção com título e borda"""
        section_frame = ctk.CTkFrame(parent, fg_color=("gray92", "gray18"))
        section_frame.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
        section_frame.grid_columnconfigure(0, weight=1)
        
        # Título da seção
        title_label = ctk.CTkLabel(
            section_frame, 
            text=title,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("gray20", "gray80")
        )
        title_label.grid(row=0, column=0, sticky="w", padx=15, pady=(10, 5))
        
        # Linha divisória
        separator = ctk.CTkFrame(section_frame, height=2, fg_color=("gray70", "gray30"))
        separator.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        
        return section_frame

    def setup_parameters_tab(self):
        tab = self.tabview.tab("Design Parameters")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame principal com scroll
        main_frame = ctk.CTkScrollableFrame(tab)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Seções de parâmetros
        sections = []
        
        # Seção 1: Parâmetros da Antena
        antenna_section = self.create_section(main_frame, "Antenna Parameters", 0, 0)
        entries = []
        row = 2
        
        def add_param(section, label, key, value, row, combo=None, check=False, tooltip=None):
            ctk.CTkLabel(
                section, 
                text=label,
                font=ctk.CTkFont(weight="bold")
            ).grid(row=row, column=0, padx=15, pady=5, sticky="w")
            
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(section, values=combo, variable=var, width=200)
                widget.grid(row=row, column=1, padx=15, pady=5)
                entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(section, text="", variable=var)
                widget.grid(row=row, column=1, padx=15, pady=5, sticky="w")
                entries.append((key, var))
            else:
                widget = ctk.CTkEntry(section, width=200)
                widget.insert(0, str(value))
                widget.grid(row=row, column=1, padx=15, pady=5)
                entries.append((key, widget))
            
            if tooltip:
                # Adicionar tooltip (implementação básica)
                pass
                
            return row + 1
        
        row = add_param(antenna_section, "Central Frequency (GHz):", "frequency", self.params["frequency"], row)
        row = add_param(antenna_section, "Desired Gain (dBi):", "gain", self.params["gain"], row)
        row = add_param(antenna_section, "Sweep Start (GHz):", "sweep_start", self.params["sweep_start"], row)
        row = add_param(antenna_section, "Sweep Stop (GHz):", "sweep_stop", self.params["sweep_stop"], row)
        row = add_param(antenna_section, "Patch Spacing:", "spacing_type", self.params["spacing_type"], row,
                       combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])
        
        # Seção 2: Parâmetros do Substrato
        substrate_section = self.create_section(main_frame, "Substrate Parameters", 1, 0)
        row = 2
        
        row = add_param(substrate_section, "Substrate Material:", "substrate_material", 
                       self.params["substrate_material"], row,
                       combo=["Rogers RO4003C (tm)", "FR4_epoxy", "Duroid (tm)", "Air"])
        row = add_param(substrate_section, "Relative Permittivity (εr):", "er", self.params["er"], row)
        row = add_param(substrate_section, "Loss Tangent (tan δ):", "tan_d", self.params["tan_d"], row)
        row = add_param(substrate_section, "Substrate Thickness (mm):", "substrate_thickness", 
                       self.params["substrate_thickness"], row)
        row = add_param(substrate_section, "Metal Thickness (mm):", "metal_thickness", 
                       self.params["metal_thickness"], row)
        
        # Seção 3: Parâmetros do Conector Coaxial
        coax_section = self.create_section(main_frame, "Coaxial Feed Parameters", 2, 0)
        row = 2
        
        row = add_param(coax_section, "Probe Radius a (mm):", "probe_radius", self.params["probe_radius"], row)
        row = add_param(coax_section, "Coax εr (PTFE):", "coax_er", self.params["coax_er"], row)
        row = add_param(coax_section, "Shield Wall (mm):", "coax_wall_thickness", 
                       self.params["coax_wall_thickness"], row)
        row = add_param(coax_section, "Port Length below GND (mm):", "coax_port_length", 
                       self.params["coax_port_length"], row)
        row = add_param(coax_section, "Anti-pad clearance (mm):", "antipad_clearance", 
                       self.params["antipad_clearance"], row)
        row = add_param(coax_section, "Feed Position:", "feed_position", self.params["feed_position"], row,
                       combo=["edge", "inset"])
        
        # Seção 4: Configurações de Simulação
        sim_section = self.create_section(main_frame, "Simulation Settings", 3, 0)
        row = 2
        
        row = add_param(sim_section, "CPU Cores:", "cores", self.params["cores"], row)
        row = add_param(sim_section, "Show HFSS Interface:", "show_gui", 
                       not self.params["non_graphical"], row, check=True)
        row = add_param(sim_section, "Save Project:", "save_project", self.save_project, row, check=True)
        
        self.entries = entries
        
        # Seção 5: Parâmetros Calculados
        calc_section = self.create_section(main_frame, "Calculated Parameters", 4, 0)
        
        # Grid para os parâmetros calculados
        calc_grid = ctk.CTkFrame(calc_section)
        calc_grid.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        
        # Colunas para organizar os parâmetros calculados
        calc_grid.columnconfigure(0, weight=1)
        calc_grid.columnconfigure(1, weight=1)
        
        self.patches_label = ctk.CTkLabel(calc_grid, text="Number of Patches: 4", 
                                         font=ctk.CTkFont(weight="bold"))
        self.patches_label.grid(row=0, column=0, sticky="w", pady=5)
        
        self.rows_cols_label = ctk.CTkLabel(calc_grid, text="Configuration: 2 x 2", 
                                           font=ctk.CTkFont(weight="bold"))
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=5)
        
        self.spacing_label = ctk.CTkLabel(calc_grid, text="Spacing: 15.0 mm (lambda/2)",
                                         font=ctk.CTkFont(weight="bold"))
        self.spacing_label.grid(row=1, column=0, sticky="w", pady=5)
        
        self.dimensions_label = ctk.CTkLabel(calc_grid, text="Patch Dimensions: 9.57 x 9.25 mm",
                                             font=ctk.CTkFont(weight="bold"))
        self.dimensions_label.grid(row=1, column=1, sticky="w", pady=5)
        
        self.lambda_label = ctk.CTkLabel(calc_grid, text="Guided Wavelength: 0.0 mm",
                                         font=ctk.CTkFont(weight="bold"))
        self.lambda_label.grid(row=2, column=0, sticky="w", pady=5)
        
        self.feed_offset_label = ctk.CTkLabel(calc_grid, text="Feed Offset: 2.0 mm",
                                              font=ctk.CTkFont(weight="bold"))
        self.feed_offset_label.grid(row=2, column=1, sticky="w", pady=5)
        
        self.substrate_dims_label = ctk.CTkLabel(calc_grid, text="Substrate Dimensions: 0.00 x 0.00 mm",
                                                 font=ctk.CTkFont(weight="bold"))
        self.substrate_dims_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=5)
        
        # Botões de ação
        button_frame = ctk.CTkFrame(calc_section)
        button_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=15)
        
        ctk.CTkButton(button_frame, text="Calculate Parameters", command=self.calculate_parameters,
                      fg_color="#2E8B57", hover_color="#3CB371", width=180).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Save Parameters", command=self.save_parameters,
                      fg_color="#4169E1", hover_color="#6495ED", width=140).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Load Parameters", command=self.load_parameters,
                      fg_color="#FF8C00", hover_color="#FFA500", width=140).pack(side="left", padx=10)

    def setup_simulation_tab(self):
        tab = self.tabview.tab("Simulation")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame principal
        main_frame = ctk.CTkFrame(tab)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        title_label = ctk.CTkLabel(main_frame, text="Simulation Control", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(pady=10)
        
        # Frame de botões
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(pady=20)
        
        self.run_button = ctk.CTkButton(button_frame, text="Run Simulation", 
                                       command=self.start_simulation_thread,
                                       fg_color="#2E8B57", hover_color="#3CB371",
                                       height=40, width=150)
        self.run_button.pack(side="left", padx=10)
        
        self.stop_button = ctk.CTkButton(button_frame, text="Stop Simulation", 
                                        command=self.stop_simulation_thread,
                                        fg_color="#DC143C", hover_color="#FF4500",
                                        height=40, width=150, state="disabled")
        self.stop_button.pack(side="left", padx=10)
        
        # Barra de progresso
        progress_frame = ctk.CTkFrame(main_frame)
        progress_frame.pack(fill="x", padx=50, pady=10)
        
        ctk.CTkLabel(progress_frame, text="Simulation Progress:", 
                    font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        
        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=20)
        self.progress_bar.pack(fill="x", pady=5)
        self.progress_bar.set(0)
        
        # Status da simulação
        self.sim_status_label = ctk.CTkLabel(main_frame, text="Simulation not started",
                                            font=ctk.CTkFont(weight="bold"))
        self.sim_status_label.pack(pady=10)
        
        # Informações adicionais
        info_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray15"))
        info_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(info_frame, 
                    text="Note: Simulation may take several minutes depending on array size and computer resources",
                    font=ctk.CTkFont(size=12, slant="italic"),
                    text_color=("gray40", "gray60")).pack(padx=10, pady=10)

    def setup_results_tab(self):
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame principal
        main_frame = ctk.CTkFrame(tab)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        
        # Título
        title_label = ctk.CTkLabel(main_frame, text="Simulation Results", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label.grid(row=0, column=0, pady=10)
        
        # Frame do gráfico
        graph_frame = ctk.CTkFrame(main_frame)
        graph_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)
        
        # Canvas para plotagem
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.fig.set_facecolor('#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF')
        self.ax.set_facecolor('#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF')
        
        # Configurar cores do gráfico para modo escuro/claro
        if ctk.get_appearance_mode() == "Dark":
            self.ax.tick_params(colors='white')
            self.ax.xaxis.label.set_color('white')
            self.ax.yaxis.label.set_color('white')
            self.ax.title.set_color('white')
            self.ax.spines['bottom'].set_color('white')
            self.ax.spines['top'].set_color('white')
            self.ax.spines['right'].set_color('white')
            self.ax.spines['left'].set_color('white')
            self.ax.grid(color='gray')
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Frame de botões de exportação
        export_frame = ctk.CTkFrame(main_frame)
        export_frame.grid(row=2, column=0, pady=10)
        
        ctk.CTkButton(export_frame, text="Export CSV", command=self.export_csv,
                      fg_color="#6A5ACD", hover_color="#7B68EE").pack(side="left", padx=10)
        ctk.CTkButton(export_frame, text="Export PNG", command=self.export_png,
                      fg_color="#20B2AA", hover_color="#40E0D0").pack(side="left", padx=10)

    def setup_log_tab(self):
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame principal
        main_frame = ctk.CTkFrame(tab)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        
        # Título
        title_label = ctk.CTkLabel(main_frame, text="Simulation Log", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label.grid(row=0, column=0, pady=10)
        
        # Área de texto para log
        self.log_text = ctk.CTkTextbox(main_frame, width=900, height=500, font=ctk.CTkFont(family="Consolas"))
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        
        # Frame de botões
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.grid(row=2, column=0, pady=10)
        
        ctk.CTkButton(button_frame, text="Clear Log", command=self.clear_log).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Save Log", command=self.save_log).pack(side="left", padx=10)

    # ------------- utilitários de log -------------
    def log_message(self, message):
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_queue.put(f"[{timestamp}] {message}\n")

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
            self.log_message(f"Error saving log: {str(e)}")

    def export_csv(self):
        try:
            if hasattr(self, 'simulation_data'):
                np.savetxt("simulation_results.csv", self.simulation_data, delimiter=",",
                           header="Frequency (GHz), S11 (dB)", comments='')
                self.log_message("Data exported to simulation_results.csv")
            else:
                self.log_message("No simulation data available for export")
        except Exception as e:
            self.log_message(f"Error exporting CSV: {str(e)}")

    def export_png(self):
        try:
            if hasattr(self, 'fig'):
                self.fig.savefig("simulation_results.png", dpi=300, bbox_inches='tight')
                self.log_message("Plot saved to simulation_results.png")
        except Exception as e:
            self.log_message(f"Error saving plot: {str(e)}")

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
                self.log_message(msg)
                return False
        self.log_message("All parameters retrieved successfully")
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
        margin = max(total_w, total_l) * 0.4  # Aumentada para 40%
        self.calculated_params["substrate_width"] = total_w + 2 * margin
        self.calculated_params["substrate_length"] = total_l + 2 * margin
        self.log_message(f"Substrate size calculated: {self.calculated_params['substrate_width']:.2f} x {self.calculated_params['substrate_length']:.2f} mm")

    def calculate_parameters(self):
        self.log_message("Starting parameter calculation")
        if not self.get_parameters():
            self.log_message("Parameter calculation failed due to invalid input")
            return
        try:
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params.update({"patch_length": L_mm, "patch_width": W_mm, "lambda_g": lambda_g_mm})
            self.calculated_params["feed_offset"] = 0.1 * W_mm

            G0 = 8.0
            desired_gain = float(self.params["gain"])
            n_est = int(math.ceil(10 ** ((desired_gain - G0) / 10.0)))
            n_est = max(1, n_est + (n_est % 2))
            rows = max(2, int(math.sqrt(n_est)))
            rows += rows % 2
            cols = max(2, int(math.ceil(n_est / rows)))
            cols += cols % 2
            while rows * cols < n_est:
                if rows <= cols:
                    rows += 2
                else:
                    cols += 2
            self.calculated_params.update({"num_patches": rows * cols, "rows": rows, "cols": cols})

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
            self.log_message("Parameters calculated successfully")
        except Exception as e:
            msg = f"Error in calculation: {str(e)}"
            self.status_label.configure(text=msg)
            self.log_message(msg)
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # --------- utilidades de modelagem ---------
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

    def _set_design_variables(self, L, W, spacing, rows, cols, h_sub, sub_w, sub_l):
        # Coax
        a = float(self.params["probe_radius"])
        er_cx = float(self.params["coax_er"])
        wall = float(self.params["coax_wall_thickness"])
        Lp = float(self.params["coax_port_length"])
        clear = float(self.params["antipad_clearance"])
        # 50Ω -> b
        b = a * math.exp(50.0 * math.sqrt(er_cx) / 60.0)

        # Variáveis globais
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
        self.hfss["eps"] = "0.001mm"         # folga numérica
        self.hfss["probeK"] = "0.3"          # fração do patchL
        self.hfss["padAir"] = f"{max(spacing, W, L)/2 + Lp + 2.0}mm"

        # Adicionar variáveis para cálculo de região de ar
        self.hfss["max_patch_dim"] = f"{max(L, W)}mm"
        self.hfss["region_offset_x"] = f"subW/2 + max_patch_dim"
        self.hfss["region_offset_y"] = f"subL/2 + max_patch_dim"
        self.hfss["region_offset_z"] = "max_patch_dim"

        return a, b, wall, Lp, clear, h_sub  # Retornar valores numéricos

    def _create_coax_feed_lumped(self, ground, substrate, x_feed: float, y_feed: float,
                                 name_prefix: str, a_val: float, b_val: float, wall_val: float, 
                                 Lp_val: float, clear_val: float, h_sub_val: float):
        """
        Constrói: Pino (a), PTFE anular (a..b) em -Lp..0, blindagem (b..b+wall) em -Lp..0,
        furo no substrato (raio b+clear) 0..h_sub, anti-pad no GND,
        e cria Lumped Port (anel) em z=-Lp.
        """
        # Usar valores numéricos passados como parâmetros
        a = a_val
        b = b_val
        wall = wall_val
        Lp = Lp_val
        clear = clear_val
        h_sub = h_sub_val

        # PINO: -Lp -> h_sub + eps
        pin = self.hfss.modeler.create_cylinder(
            orientation="Z", 
            origin=[x_feed, y_feed, -Lp], 
            radius=a,
            height=h_sub + Lp + 0.001,
            name=f"{name_prefix}_Pin", 
            material="copper"
        )

        # PTFE sólido (raio b) em -Lp..0 - manter original ao subtrair
        ptfe_solid = self.hfss.modeler.create_cylinder(
            orientation="Z", 
            origin=[x_feed, y_feed, -Lp], 
            radius=b, 
            height=Lp,
            name=f"{name_prefix}_PTFEsolid", 
            material="PTFE_Custom"
        )
        # Subtrair pino -> anel PTFE, mantendo ambos
        self.hfss.modeler.subtract(ptfe_solid, [pin], keep_originals=True)
        ptfe = self.hfss.modeler.get_object_by_name(f"{name_prefix}_PTFEsolid")
        ptfe.name = f"{name_prefix}_PTFE"

        # BLINDAGEM: tubo (b .. b+wall) em -Lp..0
        shield_outer = self.hfss.modeler.create_cylinder(
            orientation="Z", 
            origin=[x_feed, y_feed, -Lp], 
            radius=b + wall, 
            height=Lp,
            name=f"{name_prefix}_ShieldOuter", 
            material="copper"
        )
        shield_inner_void = self.hfss.modeler.create_cylinder(
            orientation="Z", 
            origin=[x_feed, y_feed, -Lp], 
            radius=b, 
            height=Lp,
            name=f"{name_prefix}_ShieldInnerVoid"
        )
        self.hfss.modeler.subtract(shield_outer, [shield_inner_void], keep_originals=True)
        
        # Renomear o shield corretamente
        shield = self.hfss.modeler.get_object_by_name(f"{name_prefix}_ShieldOuter")
        shield.name = f"{name_prefix}_Shield"

        # FURO NO SUBSTRATO (0..h_sub)
        hole_r = b + clear
        sub_hole = self.hfss.modeler.create_cylinder(
            orientation="Z", 
            origin=[x_feed, y_feed, 0.0], 
            radius=hole_r, 
            height=h_sub,
            name=f"{name_prefix}_SubHole"
        )
        self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=True)

        # ANTI-PAD NO GND (folha circular)
        g_hole = self.hfss.modeler.create_circle(
            orientation="XY", 
            origin=[x_feed, y_feed, 0.0],
            radius=hole_r, 
            name=f"{name_prefix}_GndHole"
        )
        self.hfss.modeler.subtract(ground, [g_hole], keep_originals=True)

        # PORTA LUMPED: anel entre raio a e b em z=-Lp
        port_ring = self.hfss.modeler.create_circle(
            orientation="XY", 
            origin=[x_feed, y_feed, -Lp], 
            radius=b,
            name=f"{name_prefix}_PortRing"
        )
        port_hole = self.hfss.modeler.create_circle(
            orientation="XY", 
            origin=[x_feed, y_feed, -Lp], 
            radius=a,
            name=f"{name_prefix}_PortHole"
        )
        self.hfss.modeler.subtract(port_ring, [port_hole], keep_originals=True)

        # Lumped Port - Usar a nova API do PyAEDT
        try:
            port = self.hfss.create_lumped_port_to_sheet(
                sheet=port_ring.name,
                axisdir=2,  # Direção Z
                impedance=50.0,
                portname=f"{name_prefix}_Lumped",
                renormalize=True
            )
            self.log_message(f"Lumped Port '{name_prefix}_Lumped' created successfully.")
        except Exception as e:
            self.log_message(f"Error creating lumped port: {str(e)}")
            # Tentar método alternativo se o primeiro falhar
            try:
                port = self.hfss.lumped_port(
                    assignment=port_ring.name,
                    reference=shield.name,
                    impedance=50.0,
                    name=f"{name_prefix}_Lumped",
                    renormalize=True
                )
                self.log_message(f"Lumped Port '{name_prefix}_Lumped' created with alternative method.")
            except Exception as e2:
                self.log_message(f"Error creating lumped port with alternative method: {str(e2)}")
                return pin, ptfe, None

        return pin, ptfe, shield

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

            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
            self.project_name = os.path.join(self.temp_folder.name, "patch_array.aedt")
            self.log_message(f"Creating project: {self.project_name}")
            self.progress_bar.set(0.1)

            self.log_message("Initializing HFSS")
            self.desktop = ansys.aedt.core.Desktop(
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"],
                new_desktop=True
            )
            self.progress_bar.set(0.2)

            self.log_message("Creating HFSS project")
            self.hfss = ansys.aedt.core.Hfss(
                project=self.project_name,
                design="patch_array",
                solution_type="DrivenModal",
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"]
            )
            self.log_message("HFSS initialized successfully")
            self.progress_bar.set(0.3)

            self.hfss.modeler.model_units = "mm"
            self.log_message("Model units set to: mm")

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

            # Variáveis de design
            a, b, wall, Lp, clear, h_sub_val = self._set_design_variables(L, W, spacing, rows, cols, h_sub, sub_w, sub_l)

            # Substrato e Ground
            self.log_message("Creating substrate")
            substrate = self.hfss.modeler.create_box(
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL", "h_sub"],
                name="Substrate",
                material=sub_mat
            )
            self.log_message("Creating ground plane")
            ground = self.hfss.modeler.create_rectangle(
                orientation="XY",
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL"],
                name="Ground",
                material="copper"
            )

            # Patches
            self.log_message(f"Creating {rows*cols} patches in {rows}x{cols} configuration")
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

                    # Pad de solda no patch (opcional)
                    self.hfss["probeOfsY"] = "probeK*patchL"
                    pad = self.hfss.modeler.create_circle(
                        orientation="XY", 
                        origin=[cx, f"{cy}-patchL/2+probeOfsY", "h_sub"],
                        radius="a", 
                        name=f"{patch_name}_Pad", 
                        material="copper"
                    )
                    try:
                        self.hfss.modeler.unite([patch, pad])
                    except Exception:
                        pass

                    # Coax completo + Lumped Port
                    x_feed = cx
                    y_feed = cy - L/2 + 0.3*L  # Usando valor numérico em vez de expressão
                    pin, ptfe, shield = self._create_coax_feed_lumped(
                        ground=ground, 
                        substrate=substrate,
                        x_feed=x_feed, 
                        y_feed=y_feed,
                        name_prefix=f"P{count}",
                        a_val=a,
                        b_val=b,
                        wall_val=wall,
                        Lp_val=Lp,
                        clear_val=clear,
                        h_sub_val=h_sub_val
                    )

                    if shield is None:
                        self.log_message(f"Warning: Shield for P{count} was not created correctly")

                    self.progress_bar.set(0.4 + 0.2 * (count / float(rows * cols)))

            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return

            # Região de ar + radiação (paramétrica)
            self.log_message("Creating air region + radiation boundary")
            # Usar dimensões baseadas no substrate com margem adequada
            region_size_x = f"subW/2 + {max(self.calculated_params['patch_length'], self.calculated_params['patch_width'])}"
            region_size_y = f"subL/2 + {max(self.calculated_params['patch_length'], self.calculated_params['patch_width'])}"
            region_size_z = f"{max(self.calculated_params['patch_length'], self.calculated_params['patch_width'])}"

            region = self.hfss.modeler.create_region(
                [region_size_x, region_size_y, region_size_z, region_size_x, region_size_y, region_size_z],
                is_percentage=False
            )
            self.hfss.assign_radiation_boundary_to_objects(region)
            self.progress_bar.set(0.7)

            # Setup + Sweep (201 pts)
            self.log_message("Creating simulation setup")
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02

            self.log_message("Creating frequency sweep (linear step for 201 points)")
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
                self.log_message(f"Linear-step helper not available ({e}). Using interpolating sweep.")
                setup.create_frequency_sweep(
                    unit="GHz",
                    name="Sweep1",
                    start_frequency=self.params["sweep_start"],
                    stop_frequency=self.params["sweep_stop"],
                    sweep_type="Interpolating"
                )

            # Malha leve nos patches
            self.log_message("Assigning local mesh refinement")
            try:
                lambda_g_mm = max(1e-6, self.calculated_params["lambda_g"])
                edge_len = max(lambda_g_mm / 60.0, W / 200.0)
                for p in patches:
                    self.hfss.mesh.assign_length_mesh([p], maximum_length=f"{edge_len}mm")
            except Exception as e:
                self.log_message(f"Mesh refinement warning: {e}")

            self.log_message("Validating design")
            _ = self.hfss.validate_full_design()

            self.log_message("Starting analysis")
            self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])

            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return

            self.progress_bar.set(0.9)
            self.log_message("Processing results")
            self.plot_results()
            self.progress_bar.set(1.0)
            self.sim_status_label.configure(text="Simulation completed")
            self.log_message("Simulation completed successfully")

        except Exception as e:
            msg = f"Error in simulation: {str(e)}"
            self.log_message(msg)
            self.sim_status_label.configure(text=msg)
            self.log_message(f"Traceback: {traceback.format_exc()}")
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    def plot_results(self):
        try:
            self.log_message("Plotting results")
            self.ax.clear()
            report = self.hfss.post.reports_by_category.standard(expressions=["dB(S(1,1))"])
            report.context = ["Setup1: Sweep1"]
            sol = report.get_solution_data()
            if sol:
                freqs = np.array(sol.primary_sweep_values, dtype=float)
                s11_list = sol.data_real()
                if len(s11_list) > 0:
                    s11 = np.array(s11_list[0], dtype=float)
                    self.simulation_data = np.column_stack((freqs, s11))
                    self.ax.plot(freqs, s11, label="S11", linewidth=2)
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
                    self.log_message("Results plotted successfully")
                else:
                    self.log_message("No S11 data available for plotting")
            else:
                self.log_message("Could not get simulation data")
        except Exception as e:
            self.log_message(f"Error plotting results: {str(e)}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

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
                    self.log_message(f"Error closing project: {str(e)}")
            if self.desktop and hasattr(self.desktop, 'release_desktop'):
                try:
                    self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception as e:
                    self.log_message(f"Error releasing desktop: {str(e)}")
            if self.temp_folder and not self.save_project:
                try:
                    self.temp_folder.cleanup()
                except Exception as e:
                    self.log_message(f"Error cleaning up temporary files: {str(e)}")
        except Exception as e:
            self.log_message(f"Error during cleanup: {str(e)}")

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
            self.log_message(f"Error saving parameters: {str(e)}")

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
            self.log_message(f"Error loading parameters: {str(e)}")

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
            self.log_message("Interface updated with loaded parameters")
        except Exception as e:
            self.log_message(f"Error updating interface: {str(e)}")

    def run(self):
        try:
            self.window.mainloop()
        except Exception as e:
            self.log_message(f"Unexpected error: {str(e)}")
        finally:
            self.cleanup()


if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()