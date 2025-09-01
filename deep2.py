import os
import tempfile
import time
import ansys.aedt.core
from ansys.aedt.core.modeler.advanced_cad.stackup_3d import Stackup3D
import customtkinter as ctk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np
import math
import threading
import queue
from datetime import datetime
import json
import shutil
import traceback
from typing import Dict, Any, List, Tuple, Optional

# Configuração da interface gráfica
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class PatchAntennaDesigner:
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
        
        # Parâmetros padrão
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
            "substrate_thickness": 0.5,
            "metal_thickness": 0.035,
            "er": 3.55,  # Permissividade relativa para Rogers RO4003C
            "tan_d": 0.0027  # Tangente de perdas
        }
        
        # Variáveis calculadas
        self.calculated_params = {
            "num_patches": 4,
            "spacing": 15.0,
            "patch_length": 9.57,
            "patch_width": 9.25,
            "rows": 2,
            "cols": 2,
            "lambda_g": 0.0
        }
        
        # Constantes
        self.c = 3e8
        
        self.setup_gui()
        
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        self.window.geometry("1200x900")
        
        # Configurar para fechar completamente
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Configurar grid
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)
        
        # Frame de cabeçalho
        header_frame = ctk.CTkFrame(self.window, height=80)
        header_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header_frame.grid_propagate(False)
        
        # Título
        title_label = ctk.CTkLabel(header_frame, text="Patch Antenna Array Designer", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=20)
        
        # Frame principal com abas
        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        
        # Abas
        self.tabview.add("Parameters")
        self.tabview.add("Simulation")
        self.tabview.add("Results")
        self.tabview.add("Log")
        
        # Configurar abas
        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()
        
        # Frame de status
        status_frame = ctk.CTkFrame(self.window, height=40)
        status_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        status_frame.grid_propagate(False)
        
        self.status_label = ctk.CTkLabel(status_frame, text="Ready to calculate parameters")
        self.status_label.pack(pady=10)
        
        # Iniciar thread para processar logs
        self.process_log_queue()
        
    def setup_parameters_tab(self):
        tab = self.tabview.tab("Parameters")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        
        # Frame de parâmetros
        params_frame = ctk.CTkScrollableFrame(tab)
        params_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Campos de entrada
        entries = []
        row = 0
        
        # Frequência central
        freq_label = ctk.CTkLabel(params_frame, text="Central Frequency (GHz):", 
                                 font=ctk.CTkFont(weight="bold"))
        freq_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        freq_entry = ctk.CTkEntry(params_frame, width=200)
        freq_entry.insert(0, str(self.params["frequency"]))
        freq_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("frequency", freq_entry))
        row += 1
        
        # Ganho desejado
        gain_label = ctk.CTkLabel(params_frame, text="Desired Gain (dBi):", 
                                 font=ctk.CTkFont(weight="bold"))
        gain_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        gain_entry = ctk.CTkEntry(params_frame, width=200)
        gain_entry.insert(0, str(self.params["gain"]))
        gain_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("gain", gain_entry))
        row += 1
        
        # Início do sweep
        sweep_start_label = ctk.CTkLabel(params_frame, text="Sweep Start (GHz):", 
                                        font=ctk.CTkFont(weight="bold"))
        sweep_start_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        sweep_start_entry = ctk.CTkEntry(params_frame, width=200)
        sweep_start_entry.insert(0, str(self.params["sweep_start"]))
        sweep_start_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("sweep_start", sweep_start_entry))
        row += 1
        
        # Fim do sweep
        sweep_stop_label = ctk.CTkLabel(params_frame, text="Sweep Stop (GHz):", 
                                       font=ctk.CTkFont(weight="bold"))
        sweep_stop_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        sweep_stop_entry = ctk.CTkEntry(params_frame, width=200)
        sweep_stop_entry.insert(0, str(self.params["sweep_stop"]))
        sweep_stop_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("sweep_stop", sweep_stop_entry))
        row += 1
        
        # Número de núcleos
        cores_label = ctk.CTkLabel(params_frame, text="CPU Cores:", 
                                  font=ctk.CTkFont(weight="bold"))
        cores_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        cores_entry = ctk.CTkEntry(params_frame, width=200)
        cores_entry.insert(0, str(self.params["cores"]))
        cores_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("cores", cores_entry))
        row += 1
        
        # Material do substrato
        substrate_label = ctk.CTkLabel(params_frame, text="Substrate Material:", 
                                      font=ctk.CTkFont(weight="bold"))
        substrate_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        substrate_var = ctk.StringVar(value=self.params["substrate_material"])
        substrate_combo = ctk.CTkComboBox(params_frame, 
                                         values=["Rogers RO4003C (tm)", "FR4", "Duroid (tm)", "Air"],
                                         variable=substrate_var, width=200)
        substrate_combo.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("substrate_material", substrate_var))
        row += 1
        
        # Permissividade relativa
        er_label = ctk.CTkLabel(params_frame, text="Relative Permittivity (εr):", 
                               font=ctk.CTkFont(weight="bold"))
        er_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        er_entry = ctk.CTkEntry(params_frame, width=200)
        er_entry.insert(0, str(self.params["er"]))
        er_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("er", er_entry))
        row += 1
        
        # Tangente de perdas
        tan_d_label = ctk.CTkLabel(params_frame, text="Loss Tangent (tan δ):", 
                                  font=ctk.CTkFont(weight="bold"))
        tan_d_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        tan_d_entry = ctk.CTkEntry(params_frame, width=200)
        tan_d_entry.insert(0, str(self.params["tan_d"]))
        tan_d_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("tan_d", tan_d_entry))
        row += 1
        
        # Espessura do substrato
        substrate_thickness_label = ctk.CTkLabel(params_frame, text="Substrate Thickness (mm):", 
                                                font=ctk.CTkFont(weight="bold"))
        substrate_thickness_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        substrate_thickness_entry = ctk.CTkEntry(params_frame, width=200)
        substrate_thickness_entry.insert(0, str(self.params["substrate_thickness"]))
        substrate_thickness_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("substrate_thickness", substrate_thickness_entry))
        row += 1
        
        # Espessura do metal
        metal_thickness_label = ctk.CTkLabel(params_frame, text="Metal Thickness (mm):", 
                                            font=ctk.CTkFont(weight="bold"))
        metal_thickness_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        metal_thickness_entry = ctk.CTkEntry(params_frame, width=200)
        metal_thickness_entry.insert(0, str(self.params["metal_thickness"]))
        metal_thickness_entry.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("metal_thickness", metal_thickness_entry))
        row += 1
        
        # Checkbox para interface gráfica
        gui_label = ctk.CTkLabel(params_frame, text="Show HFSS Interface:", 
                                font=ctk.CTkFont(weight="bold"))
        gui_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        gui_var = ctk.BooleanVar(value=not self.params["non_graphical"])
        gui_checkbox = ctk.CTkCheckBox(params_frame, text="", variable=gui_var)
        gui_checkbox.grid(row=row, column=1, padx=5, pady=10, sticky="w")
        entries.append(("show_gui", gui_var))
        row += 1
        
        # Opção de espaçamento
        spacing_label = ctk.CTkLabel(params_frame, text="Patch Spacing:", 
                                    font=ctk.CTkFont(weight="bold"))
        spacing_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        spacing_var = ctk.StringVar(value=self.params["spacing_type"])
        spacing_combo = ctk.CTkComboBox(params_frame, 
                                       values=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"], 
                                       variable=spacing_var, width=200)
        spacing_combo.grid(row=row, column=1, padx=5, pady=10)
        entries.append(("spacing_type", spacing_var))
        row += 1
        
        # Checkbox para salvar projeto
        save_label = ctk.CTkLabel(params_frame, text="Save Project:", 
                                 font=ctk.CTkFont(weight="bold"))
        save_label.grid(row=row, column=0, padx=5, pady=10, sticky="w")
        save_var = ctk.BooleanVar(value=self.save_project)
        save_checkbox = ctk.CTkCheckBox(params_frame, text="", variable=save_var)
        save_checkbox.grid(row=row, column=1, padx=5, pady=10, sticky="w")
        entries.append(("save_project", save_var))
        row += 1
        
        self.entries = entries
        
        # Frame de parâmetros calculados
        calc_frame = ctk.CTkFrame(tab)
        calc_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        
        calc_title = ctk.CTkLabel(calc_frame, text="Calculated Parameters", 
                                 font=ctk.CTkFont(size=16, weight="bold"))
        calc_title.pack(pady=10)
        
        # Labels para mostrar os parâmetros calculados
        self.patches_label = ctk.CTkLabel(calc_frame, text="Number of Patches: 4", 
                                         font=ctk.CTkFont(weight="bold"))
        self.patches_label.pack(pady=5)
        
        self.rows_cols_label = ctk.CTkLabel(calc_frame, text="Configuration: 2 x 2", 
                                           font=ctk.CTkFont(weight="bold"))
        self.rows_cols_label.pack(pady=5)
        
        self.spacing_label = ctk.CTkLabel(calc_frame, text="Spacing: 15.0 mm (lambda/2)", 
                                         font=ctk.CTkFont(weight="bold"))
        self.spacing_label.pack(pady=5)
        
        self.dimensions_label = ctk.CTkLabel(calc_frame, text="Patch Dimensions: 9.57 x 9.25 mm", 
                                            font=ctk.CTkFont(weight="bold"))
        self.dimensions_label.pack(pady=5)
        
        self.lambda_label = ctk.CTkLabel(calc_frame, text="Guided Wavelength: 0.0 mm", 
                                        font=ctk.CTkFont(weight="bold"))
        self.lambda_label.pack(pady=5)
        
        # Botões
        button_frame = ctk.CTkFrame(calc_frame)
        button_frame.pack(pady=20)
        
        calc_button = ctk.CTkButton(button_frame, text="Calculate Parameters", 
                                   command=self.calculate_parameters, 
                                   fg_color="green", hover_color="darkgreen")
        calc_button.pack(side="left", padx=10, pady=10)
        
        save_button = ctk.CTkButton(button_frame, text="Save Parameters", 
                                   command=self.save_parameters,
                                   fg_color="blue", hover_color="darkblue")
        save_button.pack(side="left", padx=10, pady=10)
        
        load_button = ctk.CTkButton(button_frame, text="Load Parameters", 
                                   command=self.load_parameters,
                                   fg_color="orange", hover_color="darkorange")
        load_button.pack(side="left", padx=10, pady=10)
        
    def setup_simulation_tab(self):
        tab = self.tabview.tab("Simulation")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame de simulação
        sim_frame = ctk.CTkFrame(tab)
        sim_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        sim_title = ctk.CTkLabel(sim_frame, text="Simulation Control", 
                                font=ctk.CTkFont(size=16, weight="bold"))
        sim_title.pack(pady=10)
        
        # Botões de controle de simulação
        button_frame = ctk.CTkFrame(sim_frame)
        button_frame.pack(pady=20)
        
        self.run_button = ctk.CTkButton(button_frame, text="Run Simulation", 
                                       command=self.start_simulation_thread,
                                       fg_color="green", hover_color="darkgreen")
        self.run_button.pack(side="left", padx=10, pady=10)
        
        self.stop_button = ctk.CTkButton(button_frame, text="Stop Simulation", 
                                        command=self.stop_simulation_thread,
                                        fg_color="red", hover_color="darkred",
                                        state="disabled")
        self.stop_button.pack(side="left", padx=10, pady=10)
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(sim_frame, width=400)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)
        
        # Status da simulação
        self.sim_status_label = ctk.CTkLabel(sim_frame, text="Simulation not started")
        self.sim_status_label.pack(pady=10)
        
    def setup_results_tab(self):
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame de resultados
        results_frame = ctk.CTkFrame(tab)
        results_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        results_title = ctk.CTkLabel(results_frame, text="Simulation Results", 
                                    font=ctk.CTkFont(size=16, weight="bold"))
        results_title.pack(pady=10)
        
        # Canvas para plotagem
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=results_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Botões de exportação
        export_frame = ctk.CTkFrame(results_frame)
        export_frame.pack(pady=10)
        
        export_csv_button = ctk.CTkButton(export_frame, text="Export CSV", 
                                         command=self.export_csv,
                                         fg_color="purple", hover_color="darkpurple")
        export_csv_button.pack(side="left", padx=10, pady=10)
        
        export_png_button = ctk.CTkButton(export_frame, text="Export PNG", 
                                         command=self.export_png,
                                         fg_color="teal", hover_color="darkcyan")
        export_png_button.pack(side="left", padx=10, pady=10)
        
    def setup_log_tab(self):
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame de log
        log_frame = ctk.CTkFrame(tab)
        log_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        log_title = ctk.CTkLabel(log_frame, text="Simulation Log", 
                                font=ctk.CTkFont(size=16, weight="bold"))
        log_title.pack(pady=10)
        
        # Área de texto para log
        self.log_text = ctk.CTkTextbox(log_frame, width=900, height=500)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        
        # Botões de log
        log_button_frame = ctk.CTkFrame(log_frame)
        log_button_frame.pack(pady=10)
        
        clear_log_button = ctk.CTkButton(log_button_frame, text="Clear Log", 
                                        command=self.clear_log)
        clear_log_button.pack(side="left", padx=10, pady=10)
        
        save_log_button = ctk.CTkButton(log_button_frame, text="Save Log", 
                                       command=self.save_log)
        save_log_button.pack(side="left", padx=10, pady=10)
        
    def log_message(self, message):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_queue.put(log_entry)
        
    def process_log_queue(self):
        """Processa mensagens na fila de log"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert("end", message)
                self.log_text.see("end")
        except queue.Empty:
            pass
        finally:
            # Agenda a próxima verificação se a janela ainda existe
            if self.window.winfo_exists():
                self.window.after(100, self.process_log_queue)
            
    def clear_log(self):
        """Limpa o log"""
        self.log_text.delete("1.0", "end")
        self.log_message("Log cleared")
        
    def save_log(self):
        """Salva o log em um arquivo"""
        try:
            log_content = self.log_text.get("1.0", "end")
            with open("simulation_log.txt", "w", encoding="utf-8") as f:
                f.write(log_content)
            self.log_message("Log saved to simulation_log.txt")
        except Exception as e:
            self.log_message(f"Error saving log: {str(e)}")
            
    def export_csv(self):
        """Exporta dados para CSV"""
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
        """Exporta gráfico para PNG"""
        try:
            if hasattr(self, 'fig'):
                self.fig.savefig("simulation_results.png", dpi=300, bbox_inches='tight')
                self.log_message("Plot saved to simulation_results.png")
        except Exception as e:
            self.log_message(f"Error saving plot: {str(e)}")
            
    def save_parameters(self):
        """Salva parâmetros em um arquivo JSON"""
        try:
            all_params = {**self.params, **self.calculated_params}
            with open("antenna_parameters.json", "w") as f:
                json.dump(all_params, f, indent=4)
            self.log_message("Parameters saved to antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error saving parameters: {str(e)}")
            
    def load_parameters(self):
        """Carrega parâmetros de um arquivo JSON"""
        try:
            with open("antenna_parameters.json", "r") as f:
                all_params = json.load(f)
                
            # Atualiza parâmetros
            for key in self.params:
                if key in all_params:
                    self.params[key] = all_params[key]
                    
            for key in self.calculated_params:
                if key in all_params:
                    self.calculated_params[key] = all_params[key]
                    
            # Atualiza interface
            self.update_interface_from_params()
            self.log_message("Parameters loaded from antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error loading parameters: {str(e)}")
            
    def update_interface_from_params(self):
        """Atualiza a interface com os parâmetros carregados"""
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
                        
            # Atualiza parâmetros calculados
            self.patches_label.configure(text=f"Number of Patches: {self.calculated_params['num_patches']}")
            self.rows_cols_label.configure(text=f"Configuration: {self.calculated_params['rows']} x {self.calculated_params['cols']}")
            self.spacing_label.configure(text=f"Spacing: {self.calculated_params['spacing']:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch Dimensions: {self.calculated_params['patch_length']:.2f} x {self.calculated_params['patch_width']:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {self.calculated_params['lambda_g']:.2f} mm")
            
            self.log_message("Interface updated with loaded parameters")
        except Exception as e:
            self.log_message(f"Error updating interface: {str(e)}")
        
    def get_parameters(self):
        """Obtém os parâmetros da interface"""
        self.log_message("Getting parameters from interface")
        
        for key, widget in self.entries:
            try:
                if key in ["cores"]:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = int(widget.get())
                    self.log_message(f"Parameter {key} set to: {self.params[key]}")
                elif key == "show_gui":
                    self.params["non_graphical"] = not widget.get()
                    self.log_message(f"Parameter non_graphical set to: {self.params['non_graphical']}")
                elif key == "save_project":
                    self.save_project = widget.get()
                    self.log_message(f"Parameter save_project set to: {self.save_project}")
                elif key in ["substrate_thickness", "metal_thickness", "er", "tan_d"]:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
                    self.log_message(f"Parameter {key} set to: {self.params[key]}")
                elif key == "spacing_type" or key == "substrate_material":
                    self.params[key] = widget.get()
                    self.log_message(f"Parameter {key} set to: {self.params[key]}")
                else:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
                    self.log_message(f"Parameter {key} set to: {self.params[key]}")
            except ValueError as e:
                error_msg = f"Error: Invalid value for {key}: {str(e)}"
                self.status_label.configure(text=error_msg)
                self.log_message(error_msg)
                return False
            except Exception as e:
                error_msg = f"Unexpected error getting parameter {key}: {str(e)}"
                self.status_label.configure(text=error_msg)
                self.log_message(error_msg)
                return False
                
        self.log_message("All parameters retrieved successfully")
        return True
    
    def calculate_parameters(self):
        """Calcula os parâmetros do array baseado no ganho desejado"""
        self.log_message("Starting parameter calculation")
        
        if not self.get_parameters():
            self.log_message("Parameter calculation failed due to invalid input")
            return
            
        try:
            self.log_message("Calculating patch dimensions")
            
            # Calcula as dimensões do patch
            patch_length, patch_width = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params["patch_length"] = patch_length
            self.calculated_params["patch_width"] = patch_width
            self.log_message(f"Patch dimensions calculated: {patch_length:.2f} x {patch_width:.2f} mm")
            
            # Calcula o comprimento de onda guiado
            freq = self.params["frequency"] * 1e9
            lambda0 = self.c / freq
            lambda_g = lambda0 / math.sqrt(self.params["er"])
            self.calculated_params["lambda_g"] = lambda_g * 1000  # converter para mm
            self.log_message(f"Guided wavelength calculated: {self.calculated_params['lambda_g']:.2f} mm")
            
            # Calcula o número de patches necessário para o ganho desejado
            G0 = 8.0  # Ganho aproximado de um patch individual em dBi
            desired_gain = self.params["gain"]
            
            # Calcula o número de patches necessário e garante que seja par
            num_patches = int(math.ceil(10 ** ((desired_gain - G0) / 10)))
            self.log_message(f"Initial patch count calculation: {num_patches} patches")
            
            # Garante número par de patches
            if num_patches % 2 != 0:
                num_patches += 1
                self.log_message(f"Adjusted to even number: {num_patches} patches")
                
            # Determina a melhor configuração de linhas e colunas
            rows = int(math.sqrt(num_patches))
            if rows % 2 != 0:
                rows += 1
                self.log_message(f"Adjusted rows to even number: {rows}")
                
            cols = num_patches // rows
            if cols % 2 != 0:
                cols += 1
                self.log_message(f"Adjusted columns to even number: {cols}")
                
            # Ajusta se necessário
            while rows * cols < num_patches:
                if rows <= cols:
                    rows += 2
                else:
                    cols += 2
                self.log_message(f"Adjusted configuration to {rows} x {cols}")
                    
            num_patches = rows * cols  # Número final de patches
            self.log_message(f"Final patch count: {num_patches} patches in {rows} x {cols} configuration")
            
            self.calculated_params["num_patches"] = num_patches
            self.calculated_params["rows"] = rows
            self.calculated_params["cols"] = cols
            
            # Calcula o espaçamento baseado na opção escolhida
            spacing_factor = 0.5  # padrão lambda/2
            
            if self.params["spacing_type"] == "lambda":
                spacing_factor = 1.0
            elif self.params["spacing_type"] == "0.7*lambda":
                spacing_factor = 0.7
            elif self.params["spacing_type"] == "0.8*lambda":
                spacing_factor = 0.8
            elif self.params["spacing_type"] == "0.9*lambda":
                spacing_factor = 0.9
                
            spacing = spacing_factor * self.calculated_params["lambda_g"]
            self.calculated_params["spacing"] = spacing
            self.log_message(f"Spacing calculated: {spacing:.2f} mm ({self.params['spacing_type']})")
            
            # Atualiza a interface com os valores calculados
            self.patches_label.configure(text=f"Number of Patches: {num_patches}")
            self.rows_cols_label.configure(text=f"Configuration: {rows} x {cols}")
            self.spacing_label.configure(text=f"Spacing: {spacing:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch Dimensions: {patch_length:.2f} x {patch_width:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {self.calculated_params['lambda_g']:.2f} mm")
            
            success_msg = "Parameters calculated successfully!"
            self.status_label.configure(text=success_msg)
            self.log_message(success_msg)
            self.log_message(f"Number of patches: {num_patches}")
            self.log_message(f"Array configuration: {rows} x {cols}")
            self.log_message(f"Patch dimensions: {patch_length:.2f} x {patch_width:.2f} mm")
            self.log_message(f"Spacing: {spacing:.2f} mm")
            self.log_message(f"Guided wavelength: {self.calculated_params['lambda_g']:.2f} mm")
            
        except Exception as e:
            error_msg = f"Error in calculation: {str(e)}"
            self.status_label.configure(text=error_msg)
            self.log_message(error_msg)
            self.log_message(f"Traceback: {traceback.format_exc()}")
    
    def calculate_patch_dimensions(self, frequency):
        """Calcula as dimensões do patch baseado na frequência e permissividade"""
        # Fórmula mais precisa para cálculo de patch retangular
        freq_hz = frequency * 1e9
        er = self.params["er"]
        h = self.params["substrate_thickness"] / 1000  # converter para metros
        
        # Comprimento de onda no dielétrico
        lambda0 = self.c / freq_hz
        lambda_g = lambda0 / math.sqrt(er)
        
        # Largura do patch (W)
        W = self.c / (2 * freq_hz) * math.sqrt(2 / (er + 1))
        
        # Comprimento efetivo
        eeff = (er + 1) / 2 + (er - 1) / 2 * math.pow(1 + 12 * h / W, -0.5)
        
        # Extensão do comprimento
        delta_L = 0.412 * h * (eeff + 0.3) * (W/h + 0.264) / ((eeff - 0.258) * (W/h + 0.8))
        
        # Comprimento do patch (L)
        L = lambda_g / 2 - 2 * delta_L
        
        # Converter para mm
        W_mm = W * 1000
        L_mm = L * 1000
        
        return L_mm, W_mm
    
    def start_simulation_thread(self):
        """Inicia a thread de simulação"""
        if self.is_simulation_running:
            self.log_message("Simulation is already running")
            return
            
        self.stop_simulation = False
        self.is_simulation_running = True
        self.simulation_thread = threading.Thread(target=self.run_simulation)
        self.simulation_thread.daemon = True
        self.simulation_thread.start()
        
    def stop_simulation_thread(self):
        """Para a simulação"""
        self.stop_simulation = True
        self.log_message("Simulation stop requested")
        
    def run_simulation(self):
        """Executa a simulação completa em uma thread separada"""
        try:
            self.log_message("Starting simulation")
            self.run_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            self.sim_status_label.configure(text="Simulation in progress")
            self.progress_bar.set(0)
            
            # Se os parâmetros não foram calculados, calcula agora
            if self.calculated_params["num_patches"] == 4:
                self.log_message("Parameters not calculated yet, calculating now")
                self.calculate_parameters()
                
            # Cria diretório temporário
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
            self.project_name = os.path.join(self.temp_folder.name, "patch_array.aedt")
            
            self.log_message(f"Creating project: {self.project_name}")
            self.progress_bar.set(0.1)
            
            # Inicializa HFSS
            self.log_message("Initializing HFSS")
            self.desktop = ansys.aedt.core.Desktop(
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"],
                new_desktop=True
            )
            
            self.progress_bar.set(0.2)
            
            # Cria projeto HFSS
            self.log_message("Creating HFSS project")
            self.hfss = ansys.aedt.core.Hfss(
                project=self.project_name,
                design="patch_array",
                solution_type="Terminal",
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"]
            )
            
            self.log_message("HFSS initialized successfully")
            self.progress_bar.set(0.3)
            
            # Configura unidades
            length_units = "mm"
            self.hfss.modeler.model_units = length_units
            self.log_message(f"Model units set to: {length_units}")
            
            # Cria stackup
            self.log_message("Creating stackup")
            stackup = Stackup3D(self.hfss)
            
            self.log_message("Adding ground layer")
            ground = stackup.add_ground_layer(
                "ground", material="copper", 
                thickness=self.params["metal_thickness"], 
                fill_material="air"
            )
            
            self.log_message("Adding dielectric layer")
            # Adiciona propriedades do material se não for um material padrão
            if self.params["substrate_material"] not in self.hfss.materials.material_keys:
                self.hfss.materials.add_material(
                    name=self.params["substrate_material"],
                    permittivity=self.params["er"],
                    dielectric_loss_tangent=self.params["tan_d"]
                )
                self.log_message(f"Custom material {self.params['substrate_material']} created")
            
            dielectric = stackup.add_dielectric_layer(
                "dielectric", 
                thickness=f"{self.params['substrate_thickness']}{length_units}", 
                material=self.params["substrate_material"]
            )
            
            self.log_message("Adding signal layer")
            signal = stackup.add_signal_layer(
                "signal", material="copper", 
                thickness=self.params["metal_thickness"], 
                fill_material="air"
            )
            
            self.progress_bar.set(0.4)
            
            # Obtém parâmetros calculados
            patch_length = self.calculated_params["patch_length"]
            patch_width = self.calculated_params["patch_width"]
            spacing = self.calculated_params["spacing"]
            num_patches = self.calculated_params["num_patches"]
            rows = self.calculated_params["rows"]
            cols = self.calculated_params["cols"]
            
            # Cria múltiplos patches em uma grade 2D
            self.log_message(f"Creating {num_patches} patches in {rows}x{cols} configuration")
            patches = []
            patch_count = 0
            
            # Calcula as dimensões totais do array
            total_width = (cols - 1) * (patch_width + spacing) + patch_width
            total_height = (rows - 1) * (patch_length + spacing) + patch_length
            
            # Calcula a posição inicial para centralizar o array
            start_x = -total_width / 2 + patch_width/2
            start_y = -total_height / 2 + patch_length/2
            
            for row in range(rows):
                for col in range(cols):
                    if patch_count >= num_patches or self.stop_simulation:
                        break
                        
                    self.log_message(f"Creating patch {patch_count+1} at position ({row}, {col})")
                    
                    # Cria o patch
                    patch_name = f"Patch_{patch_count+1}"
                    patch = self.hfss.modeler.create_rectangle(
                        csPlane=0, 
                        position=[start_x + col * (patch_width + spacing), 
                                 start_y + row * (patch_length + spacing), 
                                 self.params["substrate_thickness"] + self.params["metal_thickness"]/2],
                        size=[patch_width, patch_length],
                        name=patch_name,
                        matname="copper"
                    )
                    
                    # Cria porta de alimentação (lumped port)
                    port_name = f"Port_{patch_count+1}"
                    port_width = min(patch_width, patch_length) * 0.1  # 10% da menor dimensão
                    
                    # Cria retângulo para a porta
                    port_rect = self.hfss.modeler.create_rectangle(
                        csPlane=0,
                        position=[start_x + col * (patch_width + spacing) - port_width/2, 
                                 start_y + row * (patch_length + spacing) + patch_length/2 - port_width/2, 
                                 self.params["substrate_thickness"]],
                        size=[port_width, port_width],
                        name=f"PortRect_{patch_count+1}"
                    )
                    
                    # Cria porta lumped
                    port = self.hfss.create_lumped_port_to_sheet(
                        sheet=port_rect.name,
                        axisdir=0,  # Direção Z
                        impedance=50,
                        portname=port_name
                    )
                    
                    patches.append(patch)
                    patch_count += 1
                    
                    # Atualiza progresso
                    progress = 0.4 + 0.2 * (patch_count / num_patches)
                    self.progress_bar.set(progress)
            
            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return
                
            # Cria região de ar e boundary de radiação
            self.log_message("Creating radiation boundary")
            pad_length = max(total_width, total_height) * 0.5  # 50% de margem
            region = self.hfss.modeler.create_region([pad_length]*6, is_percentage=False)
            self.hfss.assign_radiation_boundary_to_objects(region)
            
            self.progress_bar.set(0.7)
            
            # Define configuração de simulação
            self.log_message("Creating simulation setup")
            setup = self.hfss.create_setup(
                name="Setup1", 
                setup_type="HFSSDriven"
            )
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02  # Critério de convergência mais rigoroso
            
            # Cria sweep de frequência
            self.log_message("Creating frequency sweep")
            sweep = setup.create_frequency_sweep(
                unit="GHz",
                name="Sweep1",
                start_frequency=self.params["sweep_start"],
                stop_frequency=self.params["sweep_stop"],
                sweep_type="Fast",
                number_of_points=201
            )
            
            self.progress_bar.set(0.8)
            
            # Valida o projeto antes de simular
            self.log_message("Validating design")
            validation = self.hfss.validate_full_design()
            if validation != 0:
                self.log_message(f"Design validation failed with code: {validation}")
                return
                
            # Executa a simulação
            self.log_message("Starting analysis")
            self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", num_cores=self.params["cores"])
            
            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return
                
            self.progress_bar.set(0.9)
            
            # Processa resultados
            self.log_message("Processing results")
            self.plot_results()
            
            self.progress_bar.set(1.0)
            self.log_message("Simulation completed successfully")
            self.sim_status_label.configure(text="Simulation completed")
            
        except Exception as e:
            error_msg = f"Error in simulation: {str(e)}"
            self.log_message(error_msg)
            self.sim_status_label.configure(text=error_msg)
            self.log_message(f"Traceback: {traceback.format_exc()}")
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False
            
    def plot_results(self):
        """Plota os resultados da simulação"""
        try:
            self.log_message("Plotting results")
            self.ax.clear()
            
            # Obtém dados S-parameter
            report = self.hfss.post.reports_by_category.standard(
                expressions=["dB(S(1,1))"]
            )
            
            # Configura o contexto do relatório
            report.context = ["Setup1: Sweep1"]
            
            # Obtém os dados da solução
            solution_data = report.get_solution_data()
            
            if solution_data:
                frequencies = solution_data.primary_sweep_values
                s11_data = solution_data.data_real()
                
                if len(s11_data) > 0:
                    # Armazena dados para exportação
                    self.simulation_data = np.column_stack((frequencies, s11_data[0]))
                    
                    # Plota S11
                    self.ax.plot(frequencies, s11_data[0], label="S11", linewidth=2)
                    
                    # Adiciona linha de -10dB como referência
                    self.ax.axhline(y=-10, color='r', linestyle='--', alpha=0.7, label='-10 dB')
                    
                    self.ax.set_xlabel("Frequency (GHz)")
                    self.ax.set_ylabel("S-Parameter (dB)")
                    self.ax.set_title("S-Parameter Results - Impedance Matching")
                    self.ax.legend()
                    self.ax.grid(True)
                    
                    # Destaca a frequência central
                    center_freq = self.params["frequency"]
                    self.ax.axvline(x=center_freq, color='g', linestyle='--', alpha=0.7)
                    self.ax.text(center_freq+0.1, self.ax.get_ylim()[1]-2, 
                                f'{center_freq} GHz', color='g')
                    
                    self.canvas.draw()
                    self.log_message("Results plotted successfully")
                else:
                    error_msg = "No S11 data available for plotting"
                    self.log_message(error_msg)
            else:
                error_msg = "Could not get simulation data"
                self.log_message(error_msg)
                
        except Exception as e:
            error_msg = f"Error plotting results: {str(e)}"
            self.log_message(error_msg)
            self.log_message(f"Traceback: {traceback.format_exc()}")
        
    def cleanup(self):
        """Limpa recursos após fechar a aplicação"""
        try:
            if self.hfss and hasattr(self.hfss, 'close_project'):
                try:
                    if self.save_project:
                        # Tenta salvar o projeto
                        self.hfss.save_project()
                        self.log_message(f"Project saved to: {self.project_name}")
                    else:
                        # Fecha sem salvar
                        self.hfss.close_project(save=False)
                        self.log_message("Project closed without saving")
                except Exception as e:
                    self.log_message(f"Error closing project: {str(e)}")
            
            # Libera o desktop se ainda existir
            if self.desktop and hasattr(self.desktop, 'release_desktop'):
                try:
                    self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                    self.log_message("Desktop released")
                except Exception as e:
                    self.log_message(f"Error releasing desktop: {str(e)}")
            
            # Limpa a pasta temporária se não quisermos salvar
            if self.temp_folder and not self.save_project:
                try:
                    self.temp_folder.cleanup()
                    self.log_message("Temporary files cleaned up")
                except Exception as e:
                    self.log_message(f"Error cleaning up temporary files: {str(e)}")
                    
        except Exception as e:
            self.log_message(f"Error during cleanup: {str(e)}")
    
    def on_closing(self):
        """Função chamada quando a janela é fechada"""
        self.log_message("Application closing...")
        self.cleanup()
        self.window.quit()
        self.window.destroy()
    
    def run(self):
        """Executa a aplicação"""
        try:
            self.window.mainloop()
        except Exception as e:
            self.log_message(f"Unexpected error: {str(e)}")
        finally:
            self.cleanup()

# Executa a aplicação
if __name__ == "__main__":
    app = PatchAntennaDesigner()
    app.run()