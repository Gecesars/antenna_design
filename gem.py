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
from typing import Dict, Any, List, Tuple, Optional

# Configuração da interface gráfica
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class PatchAntennaDesigner:
    """
    Uma aplicação de GUI para projetar, simular e analisar arranjos de antenas patch de microfita
    usando a biblioteca PyAEDT para automatizar o Ansys HFSS.
    """
    def __init__(self):
        self.hfss: Optional[ansys.aedt.core.Hfss] = None
        self.desktop: Optional[ansys.aedt.core.Desktop] = None
        self.temp_folder: Optional[tempfile.TemporaryDirectory] = None
        self.project_name: str = ""
        self.log_queue: queue.Queue = queue.Queue()
        self.simulation_thread: Optional[threading.Thread] = None
        self.stop_simulation: bool = False
        self.save_project: bool = False
        self.is_simulation_running: bool = False
        self.closing: bool = False
        
        self.params: Dict[str, Any] = {
            "frequency": 10.0, "gain": 12.0, "sweep_start": 8.0, "sweep_stop": 12.0,
            "cores": 4, "aedt_version": "2024.2", "non_graphical": False,
            "spacing_type": "0.7*lambda", "substrate_material": "Rogers RO4003C (tm)",
            "substrate_thickness": 0.5, "metal_thickness": 0.035, "er": 3.55,
            "tan_d": 0.0027, "feed_position": "inset"
        }
        
        self.calculated_params: Dict[str, Any] = {
            "num_patches": 4, "spacing": 15.0, "patch_length": 9.57, "patch_width": 9.25,
            "rows": 2, "cols": 2, "lambda_g": 0.0, "inset_distance": 2.0, "inset_gap": 0.5,
            "substrate_width": 50.0, "substrate_length": 50.0, "feed_line_width_50": 1.15,
            "feed_line_width_70": 0.7, "feed_line_width_100": 0.35,
        }
        
        self.simulation_results = {}
        self.c: float = 299792458.0
        self.setup_gui()
        
    def setup_gui(self):
        """Configura a interface gráfica principal usando customtkinter."""
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        self.window.geometry("1300x950")
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.grid_columnconfigure(0, weight=1); self.window.grid_rowconfigure(1, weight=1)
        
        header_frame = ctk.CTkFrame(self.window, height=80, corner_radius=10)
        header_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header_frame.pack_propagate(False)
        ctk.CTkLabel(header_frame, text="Patch Antenna Array Designer", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=20)
        
        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.tabview.add("Parameters"); self.tabview.add("Simulation"); self.tabview.add("Results"); self.tabview.add("Log")
        
        self.setup_parameters_tab(); self.setup_simulation_tab(); self.setup_results_tab(); self.setup_log_tab()
        
        status_frame = ctk.CTkFrame(self.window, height=40, corner_radius=10)
        status_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        status_frame.pack_propagate(False)
        self.status_label = ctk.CTkLabel(status_frame, text="Ready. Please calculate parameters to begin.")
        self.status_label.pack(pady=10)
        
        self.process_log_queue()

    def setup_parameters_tab(self):
        tab = self.tabview.tab("Parameters")
        tab.grid_columnconfigure(0, weight=1); tab.grid_columnconfigure(1, weight=1); tab.grid_rowconfigure(0, weight=1)

        params_frame = ctk.CTkScrollableFrame(tab, label_text="Input Parameters")
        params_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.entries: Dict[str, Any] = {}
        row = 0
        param_map = {
            "frequency": {"label": "Central Frequency (GHz)", "type": "entry"}, "gain": {"label": "Target Gain (dBi)", "type": "entry"},
            "sweep_start": {"label": "Sweep Start (GHz)", "type": "entry"}, "sweep_stop": {"label": "Sweep Stop (GHz)", "type": "entry"},
            "cores": {"label": "CPU Cores", "type": "entry"},
            "substrate_material": {"label": "Substrate Material", "type": "combo", "values": ["Rogers RO4003C (tm)", "FR-4", "Duroid 5880 (tm)", "Teflon (tm)", "Air"]},
            "er": {"label": "Relative Permittivity (εr)", "type": "entry"}, "tan_d": {"label": "Loss Tangent (tan δ)", "type": "entry"},
            "substrate_thickness": {"label": "Substrate Thickness (mm)", "type": "entry"}, "metal_thickness": {"label": "Metal Thickness (mm)", "type": "entry"},
            "feed_position": {"label": "Feed Position", "type": "combo", "values": ["inset", "edge"]},
            "spacing_type": {"label": "Patch Spacing", "type": "combo", "values": ["0.6*lambda", "0.7*lambda", "0.8*lambda", "lambda/2", "lambda"]},
        }
        for key, config in param_map.items():
            ctk.CTkLabel(params_frame, text=config["label"], font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, padx=10, pady=8, sticky="w")
            widget = ctk.CTkEntry(params_frame, width=200) if config["type"] == "entry" else ctk.CTkComboBox(params_frame, values=config["values"], width=200)
            if isinstance(widget, ctk.CTkEntry): widget.insert(0, str(self.params[key]))
            else: widget.set(self.params[key])
            widget.grid(row=row, column=1, padx=10, pady=8, sticky="w")
            self.entries[key] = widget; row += 1
        ctk.CTkLabel(params_frame, text="Show HFSS Interface:", font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, padx=10, pady=8, sticky="w")
        self.entries["show_gui"] = ctk.CTkCheckBox(params_frame, text="", onvalue=True, offvalue=False)
        self.entries["show_gui"].select() if not self.params["non_graphical"] else self.entries["show_gui"].deselect()
        self.entries["show_gui"].grid(row=row, column=1, padx=10, pady=8, sticky="w"); row += 1
        ctk.CTkLabel(params_frame, text="Save Project on Exit:", font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, padx=10, pady=8, sticky="w")
        self.entries["save_project"] = ctk.CTkCheckBox(params_frame, text="", onvalue=True, offvalue=False)
        self.entries["save_project"].select() if self.save_project else self.entries["save_project"].deselect()
        self.entries["save_project"].grid(row=row, column=1, padx=10, pady=8, sticky="w")

        calc_frame = ctk.CTkFrame(tab, corner_radius=10); calc_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        ctk.CTkLabel(calc_frame, text="Calculated & Design Parameters", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10, anchor="n")
        self.calc_labels = {
            "patches": ctk.CTkLabel(calc_frame, text=""), "rows_cols": ctk.CTkLabel(calc_frame, text=""), "spacing": ctk.CTkLabel(calc_frame, text=""),
            "dimensions": ctk.CTkLabel(calc_frame, text=""), "lambda_g": ctk.CTkLabel(calc_frame, text=""), "inset": ctk.CTkLabel(calc_frame, text=""),
            "substrate_dims": ctk.CTkLabel(calc_frame, text=""), "feed_width_50": ctk.CTkLabel(calc_frame, text=""),
            "feed_width_70": ctk.CTkLabel(calc_frame, text=""), "feed_width_100": ctk.CTkLabel(calc_frame, text=""),
        }
        for label in self.calc_labels.values(): label.pack(pady=5, padx=10, anchor="w")
        button_frame = ctk.CTkFrame(calc_frame); button_frame.pack(pady=20, anchor="s", expand=True)
        ctk.CTkButton(button_frame, text="Calculate Parameters", command=self.calculate_parameters, fg_color="green", hover_color="darkgreen").pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Save", command=self.save_parameters).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Load", command=self.load_parameters).pack(side="left", padx=10)
        self.update_calculated_params_display()
        
    def setup_simulation_tab(self):
        tab = self.tabview.tab("Simulation"); tab.grid_columnconfigure(0, weight=1); tab.grid_rowconfigure(0, weight=1)
        sim_frame = ctk.CTkFrame(tab); sim_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        ctk.CTkLabel(sim_frame, text="Simulation Control", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        button_frame = ctk.CTkFrame(sim_frame); button_frame.pack(pady=20)
        self.run_button = ctk.CTkButton(button_frame, text="Run Simulation", command=self.start_simulation_thread, fg_color="green", hover_color="darkgreen")
        self.run_button.pack(side="left", padx=10, pady=10)
        self.stop_button = ctk.CTkButton(button_frame, text="Stop Simulation", command=self.stop_simulation_thread, fg_color="red", hover_color="darkred", state="disabled")
        self.stop_button.pack(side="left", padx=10, pady=10)
        self.progress_bar = ctk.CTkProgressBar(sim_frame, width=400); self.progress_bar.pack(pady=10); self.progress_bar.set(0)
        self.sim_status_label = ctk.CTkLabel(sim_frame, text="Simulation not started"); self.sim_status_label.pack(pady=10)

    def setup_results_tab(self):
        tab = self.tabview.tab("Results"); tab.grid_columnconfigure(0, weight=1); tab.grid_rowconfigure(1, weight=1)
        results_summary_frame = ctk.CTkFrame(tab); results_summary_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.result_labels = {
            "gain": ctk.CTkLabel(results_summary_frame, text="Simulated Gain (dBi): --", font=ctk.CTkFont(size=14, weight="bold")),
            "directivity": ctk.CTkLabel(results_summary_frame, text="Simulated Directivity (dBi): --", font=ctk.CTkFont(size=14, weight="bold")),
            "min_s11": ctk.CTkLabel(results_summary_frame, text="Min S11 (dB): --", font=ctk.CTkFont(size=14, weight="bold"))
        }
        for label in self.result_labels.values(): label.pack(side="left", padx=20, pady=10)
        plot_frame = ctk.CTkFrame(tab); plot_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.fig, self.ax = plt.subplots(figsize=(8, 6), facecolor="#2B2B2B")
        self.ax.tick_params(axis='x', colors='white'); self.ax.tick_params(axis='y', colors='white')
        self.ax.xaxis.label.set_color('white'); self.ax.yaxis.label.set_color('white'); self.ax.title.set_color('white')
        self.ax.set_facecolor("#3C3C3C")
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        export_frame = ctk.CTkFrame(tab); export_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        ctk.CTkButton(export_frame, text="Export CSV", command=self.export_csv).pack(side="left", padx=10, pady=5)
        ctk.CTkButton(export_frame, text="Export PNG", command=self.export_png).pack(side="left", padx=10, pady=5)
        
    def setup_log_tab(self):
        tab = self.tabview.tab("Log"); tab.grid_columnconfigure(0, weight=1); tab.grid_rowconfigure(0, weight=1)
        log_frame = ctk.CTkFrame(tab); log_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        ctk.CTkLabel(log_frame, text="Simulation Log", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        self.log_text = ctk.CTkTextbox(log_frame, width=900, height=500)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        log_button_frame = ctk.CTkFrame(log_frame); log_button_frame.pack(pady=10)
        ctk.CTkButton(log_button_frame, text="Clear Log", command=self.clear_log).pack(side="left", padx=10, pady=10)
        ctk.CTkButton(log_button_frame, text="Save Log", command=self.save_log).pack(side="left", padx=10, pady=10)
        
    def get_parameters(self) -> bool:
        self.log_message("Getting parameters from interface...")
        try:
            for key, widget in self.entries.items():
                value = widget.get()
                if key in ["frequency", "gain", "sweep_start", "sweep_stop", "er", "tan_d", "substrate_thickness", "metal_thickness"]: self.params[key] = float(value)
                elif key == "cores": self.params[key] = int(value)
                elif key == "show_gui": self.params["non_graphical"] = not value
                elif key == "save_project": self.save_project = value
                else: self.params[key] = value
            self.log_message("Parameters retrieved successfully."); return True
        except (ValueError, TypeError) as e:
            msg = f"Invalid input value: {e}"; self.log_message(f"ERROR: {msg}"); self.status_label.configure(text=msg); return False

    def calculate_microstrip_width(self, Z0: float, er: float, h: float) -> float:
        A = (Z0 / 60.0) * math.sqrt((er + 1.0) / 2.0) + ((er - 1.0) / (er + 1.0)) * (0.23 + 0.11 / er)
        W_h_ratio = (8.0 * math.exp(A)) / (math.exp(2.0 * A) - 2.0)
        if W_h_ratio >= 2.0:
            B = (377.0 * math.pi) / (2.0 * Z0 * math.sqrt(er))
            W_h_ratio = (2.0 / math.pi) * (B - 1.0 - math.log(2.0 * B - 1.0) + ((er - 1.0) / (2.0 * er)) * (math.log(B - 1.0) + 0.39 - (0.61 / er)))
        return W_h_ratio * h

    def calculate_parameters(self):
        if not self.get_parameters(): return
        self.log_message("Starting antenna parameter calculation...")
        try:
            freq_hz, er, h_m = self.params["frequency"] * 1e9, self.params["er"], self.params["substrate_thickness"] / 1000.0
            W_m = (self.c / (2 * freq_hz)) * math.sqrt(2 / (er + 1))
            self.calculated_params["patch_width"] = W_m * 1000.0
            ereff = (er + 1) / 2 + (er - 1) / 2 * math.pow(1 + 12 * h_m / W_m, -0.5)
            delta_L_m = 0.412 * h_m * ((ereff + 0.3) * (W_m / h_m + 0.264)) / ((ereff - 0.258) * (W_m / h_m + 0.8))
            lambda_eff_m = self.c / (freq_hz * math.sqrt(ereff))
            L_m = lambda_eff_m / 2 - 2 * delta_L_m
            self.calculated_params["patch_length"] = L_m * 1000.0
            self.calculated_params["lambda_g"] = (self.c / (freq_hz * math.sqrt(er))) * 1000.0
            try:
                inset_dist_m = (L_m / math.pi) * math.acos(math.sqrt(50.0 / 240.0))
                self.calculated_params["inset_distance"] = inset_dist_m * 1000.0
            except ValueError: self.calculated_params["inset_distance"] = L_m / 3.0 * 1000.0
            w50 = self.calculate_microstrip_width(50.0, er, h_m * 1000.0)
            self.calculated_params["inset_gap"] = w50
            gain_single_patch_dbi, target_gain_linear = 6.5, 10**(self.params["gain"] / 10.0)
            single_gain_linear = 10**(gain_single_patch_dbi / 10.0)
            num_patches_req = max(1.0, target_gain_linear / single_gain_linear)
            rows = round(math.sqrt(num_patches_req)); cols = round(num_patches_req / rows) if rows > 0 else 0
            num_patches = max(1, rows * cols)
            self.calculated_params.update({"num_patches": num_patches, "rows": rows, "cols": cols})
            spacing_type = self.params["spacing_type"]
            if "/" in spacing_type: parts = spacing_type.replace("lambda", "1.0").split('/'); spacing_factor = float(parts[0]) / float(parts[1])
            elif "*" in spacing_type: parts = spacing_type.split('*'); spacing_factor = float(parts[0])
            else: spacing_factor = 0.5
            self.calculated_params["spacing"] = spacing_factor * (self.c / freq_hz * 1000.0)
            self.calculate_substrate_size()
            self.calculated_params["feed_line_width_50"] = w50
            self.calculated_params["feed_line_width_70"] = self.calculate_microstrip_width(70.7, er, h_m * 1000.0)
            self.calculated_params["feed_line_width_100"] = self.calculate_microstrip_width(100.0, er, h_m * 1000.0)
            self.update_calculated_params_display()
            self.status_label.configure(text="Parameters calculated successfully. Ready for simulation.")
            self.log_message("Parameter calculation finished successfully.")
        except Exception as e:
            msg = f"Error during calculation: {e}"; self.log_message(f"ERROR: {msg}\n{traceback.format_exc()}"); self.status_label.configure(text=msg)

    def calculate_substrate_size(self):
        cp = self.calculated_params
        total_width = cp["cols"] * cp["patch_width"] + (cp["cols"] - 1) * cp["spacing"]
        total_length = cp["rows"] * cp["patch_length"] + (cp["rows"] - 1) * cp["spacing"]
        lambda0_mm = (self.c / (self.params["frequency"] * 1e9)) * 1000.0
        margin = max(max(total_width, total_length) * 0.20, lambda0_mm / 4.0)
        cp["substrate_width"] = total_width + 2 * margin
        cp["substrate_length"] = total_length + 2 * margin
        self.log_message(f"Substrate size calculated: {cp['substrate_width']:.2f} x {cp['substrate_length']:.2f} mm")

    def update_calculated_params_display(self):
        cp = self.calculated_params
        self.calc_labels["patches"].configure(text=f"Number of Patches: {cp['num_patches']}")
        self.calc_labels["rows_cols"].configure(text=f"Configuration: {cp['rows']} x {cp['cols']}")
        self.calc_labels["spacing"].configure(text=f"Spacing: {cp['spacing']:.2f} mm ({self.params['spacing_type']})")
        self.calc_labels["dimensions"].configure(text=f"Patch Dimensions (LxW): {cp['patch_length']:.2f} x {cp['patch_width']:.2f} mm")
        self.calc_labels["lambda_g"].configure(text=f"Guided Wavelength (λg): {cp['lambda_g']:.2f} mm")
        self.calc_labels["inset"].configure(text=f"Inset Feed (Dist x Gap): {cp['inset_distance']:.2f} x {cp['inset_gap']:.2f} mm")
        self.calc_labels["substrate_dims"].configure(text=f"Substrate (LxW): {cp['substrate_length']:.2f} x {cp['substrate_width']:.2f} mm")
        self.calc_labels["feed_width_50"].configure(text=f"50Ω Line Width: {cp['feed_line_width_50']:.2f} mm")
        self.calc_labels["feed_width_70"].configure(text=f"70.7Ω Line Width: {cp['feed_line_width_70']:.2f} mm")
        self.calc_labels["feed_width_100"].configure(text=f"100Ω Line Width: {cp['feed_line_width_100']:.2f} mm")
        
    def _ensure_material_exists(self, material_name: str):
        """Verifica se um material existe no HFSS e o cria se necessário."""
        # **CORREÇÃO FINAL: Usa 'material_keys' e define propriedades no objeto material**
        if material_name not in self.hfss.materials.material_keys:
            self.log_message(f"Material '{material_name}' not found. Creating it programmatically.")
            try:
                new_mat = self.hfss.materials.add_material(material_name)
                new_mat.permittivity = self.params["er"]
                new_mat.dielectric_loss_tangent = self.params["tan_d"]
                new_mat.update()
                self.log_message(f"Material '{material_name}' created successfully.")
            except Exception as e:
                self.log_message(f"Failed to create material '{material_name}': {e}")
                raise

    def run_simulation(self):
        try:
            self.log_message("Simulation thread started.")
            self.run_button.configure(state="disabled"); self.stop_button.configure(state="normal")
            self.sim_status_label.configure(text="Initializing AEDT..."); self.progress_bar.set(0)
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
            self.project_name = os.path.join(self.temp_folder.name, "PatchArrayProject.aedt")
            with ansys.aedt.core.Desktop(self.params["aedt_version"], self.params["non_graphical"], new_desktop=True) as self.desktop:
                self.hfss = ansys.aedt.core.Hfss(project=self.project_name, design="PatchArray_HFSS")
                self.log_message(f"AEDT {self.params['aedt_version']} initialized.")
                self.hfss.modeler.model_units = "mm"
                self.sim_status_label.configure(text="Creating 3D Model..."); self.progress_bar.set(0.1)
                self.create_geometry()
                if self.stop_simulation: return
                self.sim_status_label.configure(text="Configuring Analysis..."); self.progress_bar.set(0.6)
                lambda0_mm = (self.c / (self.params["frequency"] * 1e9)) * 1000.0
                self.hfss.modeler.create_air_region(x_pos=lambda0_mm/2, y_pos=lambda0_mm/2, z_pos=lambda0_mm/2, x_neg=lambda0_mm/2, y_neg=lambda0_mm/2, z_neg=lambda0_mm/2)
                setup = self.hfss.create_setup("Setup1")
                setup.props["Frequency"] = f"{self.params['frequency']}GHz"
                setup.props["MaximumPasses"] = 15; setup.props["MinimumConvergedPasses"] = 2
                sweep = setup.create_frequency_sweep(unit="GHz", name="Sweep1", start_frequency=self.params["sweep_start"], stop_frequency=self.params["sweep_stop"], sweep_type="Interpolating")
                sweep.props["SaveFields"] = False; sweep.props["GenerateSurfaceCurrent"] = True
                self.sim_status_label.configure(text="Running Simulation..."); self.progress_bar.set(0.7)
                # **CORREÇÃO DO ARGUMENTO DEPRECIADO**
                self.hfss.analyze_setup("Setup1", cores=self.params["cores"])
                if self.stop_simulation: return
                self.sim_status_label.configure(text="Processing Results..."); self.progress_bar.set(0.9)
                self.post_process_results()
                self.sim_status_label.configure(text="Simulation Completed."); self.progress_bar.set(1.0)
        except Exception as e:
            msg = f"Simulation failed: {e}"; self.log_message(f"FATAL ERROR: {msg}\n{traceback.format_exc()}"); self.sim_status_label.configure(text=msg)
        finally:
            self.cleanup()
            self.run_button.configure(state="normal"); self.stop_button.configure(state="disabled")
            self.is_simulation_running = False
            self.log_message("Simulation thread finished.")

    def create_geometry(self):
        cp, p, z_pos = self.calculated_params, self.params, self.params["substrate_thickness"]
        self._ensure_material_exists(p["substrate_material"])
        gnd = self.hfss.modeler.create_rectangle("XY", [-cp["substrate_width"]/2, -cp["substrate_length"]/2, 0], [cp["substrate_width"], cp["substrate_length"]], name="Ground", material="copper")
        sub = self.hfss.modeler.create_box([-cp["substrate_width"]/2, -cp["substrate_length"]/2, 0], sizes=[cp["substrate_width"], cp["substrate_length"], p["substrate_thickness"]], name="Substrate", material=p["substrate_material"])
        sub.transparent = 0.7; self.hfss.assign_perfecte_to_sheets(gnd)
        start_x, start_y = - (cp["cols"] - 1) * cp["spacing"] / 2, - (cp["rows"] - 1) * cp["spacing"] / 2
        patch_positions = [[start_x + c * cp["spacing"], start_y + r * cp["spacing"]] for r in range(cp["rows"]) for c in range(cp["cols"])]
        patch_objects, feed_points = [], []
        for i, (pos_x, pos_y) in enumerate(patch_positions):
            patch = self.hfss.modeler.create_rectangle("XY", [pos_x - cp["patch_width"]/2, pos_y - cp["patch_length"]/2, z_pos], [cp["patch_width"], cp["patch_length"]], name=f"Patch_{i+1}", material="copper")
            if p["feed_position"] == "inset":
                inset_notch = self.hfss.modeler.create_rectangle("XY", [pos_x - cp["inset_gap"]/2, pos_y - cp["patch_length"]/2, z_pos], [cp["inset_gap"], cp["inset_distance"]], name=f"InsetNotch_{i+1}")
                self.hfss.modeler.subtract(patch, inset_notch)
                feed_points.append([pos_x, pos_y - cp["patch_length"]/2 + cp["inset_distance"], z_pos])
            else: feed_points.append([pos_x, pos_y - cp["patch_length"]/2, z_pos])
            patch_objects.append(patch.name)
        feed_network_parts = self.create_corporate_feed(feed_points)
        self.hfss.modeler.unite(patch_objects + feed_network_parts)
        port_pos = [-cp["feed_line_width_50"]/2, -cp["substrate_length"]/2 + 0.1, 0]
        port_rect = self.hfss.modeler.create_rectangle("YZ", port_pos, [z_pos, cp["feed_line_width_50"]], name="Port1_rect")
        self.hfss.lumped_port(port_rect.name, reference=gnd.name, impedance=50, name="Port1", renormalize=True)
        
    def create_corporate_feed(self, element_feed_points: List[List[float]]) -> List[str]:
        self.log_message("Creating corporate feed network...")
        cp, p, z_pos = self.calculated_params, self.params, self.params["substrate_thickness"]
        feed_parts = []; current_level_points = element_feed_points
        w50, w70 = cp["feed_line_width_50"], cp["feed_line_width_70"]
        lambda_q = cp["lambda_g"] / 4.0
        if cp['cols'] > 1:
            next_level_points = []
            points_by_y = {};
            for pt in current_level_points:
                y = round(pt[1], 2); points_by_y.setdefault(y, []).append(pt)
            for y_coord in sorted(points_by_y.keys()):
                points = sorted(points_by_y[y_coord], key=lambda x: x[0])
                while len(points) > 1:
                    new_points_this_row = []
                    for i in range(len(points) // 2):
                        p1, p2 = points[i*2], points[i*2 + 1]
                        mid_x = (p1[0] + p2[0]) / 2.0
                        feed_parts.append(self.hfss.modeler.create_rectangle("XY", [p1[0], p1[1] - w50/2, z_pos], [mid_x - p1[0], w50]).name)
                        feed_parts.append(self.hfss.modeler.create_rectangle("XY", [p2[0], p2[1] - w50/2, z_pos], [mid_x - p2[0], w50]).name)
                        trans = self.hfss.modeler.create_rectangle("XY", [mid_x - w70/2, p1[1], z_pos], [w70, -lambda_q])
                        feed_parts.append(trans.name)
                        new_points_this_row.append([mid_x, p1[1] - lambda_q, z_pos])
                    if len(points) % 2 != 0: new_points_this_row.append(points[-1])
                    points = new_points_this_row
                next_level_points.extend(points)
            current_level_points = next_level_points
        while len(current_level_points) > 1:
            next_level_points = []
            points = sorted(current_level_points, key=lambda p:p[1])
            for i in range(len(points) // 2):
                p1, p2 = points[i*2], points[i*2 + 1]
                mid_y = (p1[1] + p2[1]) / 2.0
                feed_parts.append(self.hfss.modeler.create_rectangle("XY", [p1[0] - w50/2, p1[1], z_pos], [w50, mid_y - p1[1]]).name)
                feed_parts.append(self.hfss.modeler.create_rectangle("XY", [p2[0] - w50/2, p2[1], z_pos], [w50, mid_y - p2[1]]).name)
                trans = self.hfss.modeler.create_rectangle("XY", [p1[0] - w70/2, mid_y, z_pos], [w70, -lambda_q])
                feed_parts.append(trans.name)
                next_level_points.append([p1[0], mid_y - lambda_q, z_pos])
            if len(points) % 2 != 0: next_level_points.append(points[-1])
            current_level_points = next_level_points
        final_point = current_level_points[0]
        main_feed_len = final_point[1] - (-cp["substrate_length"]/2 + 0.1)
        main_feed = self.hfss.modeler.create_rectangle("XY", [final_point[0] - w50/2, final_point[1], z_pos], [w50, -main_feed_len])
        feed_parts.append(main_feed.name)
        return feed_parts

    def post_process_results(self):
        try:
            report = self.hfss.post.reports_by_category.standard("dB(S(Port1,Port1))")
            solution_data = report.get_solution_data()
            if solution_data:
                freq, s11_db = np.array(solution_data.primary_sweep_values), np.array(solution_data.data_real()[0])
                self.simulation_results["s_params"] = np.column_stack((freq, s11_db))
                self.result_labels["min_s11"].configure(text=f"Min S11 (dB): {min(s11_db):.2f}")
                self.plot_s_params()
        except Exception as e: self.log_message(f"Could not extract S-parameters: {e}")
        try:
            farfield_setup = f"Setup1 : Freq='{self.params['frequency']}GHz' Phase='0deg'"
            report_gain = self.hfss.post.reports_by_category.far_field("RealizedGain", "dB10"); report_gain.far_field_sphere_setup = farfield_setup
            gain_data = report_gain.get_solution_data()
            if gain_data: self.simulation_results["gain"] = max(gain_data.data_real()); self.result_labels["gain"].configure(text=f"Simulated Gain (dBi): {self.simulation_results['gain']:.2f}")
            report_dir = self.hfss.post.reports_by_category.far_field("Directivity", "dBi"); report_dir.far_field_sphere_setup = farfield_setup
            dir_data = report_dir.get_solution_data()
            if dir_data: self.simulation_results["directivity"] = max(dir_data.data_real()); self.result_labels["directivity"].configure(text=f"Simulated Directivity (dBi): {self.simulation_results['directivity']:.2f}")
        except Exception as e: self.log_message(f"Could not extract Far Field data: {e}\n{traceback.format_exc()}")

    def plot_s_params(self):
        if "s_params" in self.simulation_results:
            data = self.simulation_results["s_params"]
            self.ax.clear()
            self.ax.plot(data[:, 0], data[:, 1], label="S11", linewidth=2, color="#1f77b4")
            self.ax.axhline(y=-10, color='red', linestyle='--', alpha=0.8, label='-10 dB (Matching Ref.)')
            self.ax.set_xlabel("Frequency (GHz)"); self.ax.set_ylabel("S-Parameter (dB)"); self.ax.set_title("S11 - Return Loss")
            self.ax.legend(facecolor='#4d4d4d', edgecolor='white', labelcolor='white')
            self.ax.grid(True, linestyle='--', alpha=0.4)
            self.ax.set_ylim(min(data[:, 1]) - 5 if len(data[:, 1]) > 0 else -40, 0)
            self.canvas.draw()
            self.log_message("S-parameter plot updated.")

    def on_closing(self):
        self.log_message("Application closing...")
        self.closing = True
        if self.is_simulation_running: self.stop_simulation = True; self.log_message("Waiting for simulation thread to finish...")
        self.window.destroy()

    def cleanup(self):
        if self.hfss:
            try: self.hfss.close_project(save=self.save_project); self.log_message(f"Project closed. Saved: {self.save_project}")
            except Exception as e: self.log_message(f"Error closing project: {e}")
        if self.temp_folder and not self.save_project:
            try: self.temp_folder.cleanup(); self.log_message("Temporary files cleaned up.")
            except Exception as e: self.log_message(f"Error cleaning up temp folder: {e}")

    def run(self): self.window.mainloop()
    def export_csv(self):
        if "s_params" in self.simulation_results:
            try: np.savetxt("s11_results.csv", self.simulation_results["s_params"], delimiter=",", header="Frequency_GHz,S11_dB", comments=""); self.log_message("S-parameter data exported to s11_results.csv")
            except Exception as e: self.log_message(f"Error exporting CSV: {e}")
        else: self.log_message("No simulation data to export.")
    def export_png(self):
        try: self.fig.savefig("s11_plot.png", dpi=300, bbox_inches='tight', facecolor=self.fig.get_facecolor()); self.log_message("Plot saved to s11_plot.png")
        except Exception as e: self.log_message(f"Error saving plot: {e}")
    def save_parameters(self):
        try: self.get_parameters(); f=open("antenna_design.json","w"); json.dump({"input":self.params,"calculated":self.calculated_params},f,indent=4); f.close(); self.log_message("Design parameters saved to antenna_design.json")
        except Exception as e: self.log_message(f"Error saving parameters: {e}")
    def load_parameters(self):
        try:
            with open("antenna_design.json", "r") as f: all_params = json.load(f)
            self.params = all_params.get("input", self.params); self.calculated_params = all_params.get("calculated", self.calculated_params)
            self.update_interface_from_params(); self.log_message("Parameters loaded and interface updated.")
        except FileNotFoundError: self.log_message("Error: antenna_design.json not found.")
        except Exception as e: self.log_message(f"Error loading parameters: {e}")
    def update_interface_from_params(self):
        for key, widget in self.entries.items():
            if key=="show_gui": widget.select() if not self.params["non_graphical"] else widget.deselect()
            elif key=="save_project": widget.select() if self.save_project else widget.deselect()
            elif key in self.params:
                if isinstance(widget,ctk.CTkEntry): widget.delete(0,"end"); widget.insert(0,str(self.params[key]))
                else: widget.set(str(self.params[key]))
        self.update_calculated_params_display()
    def log_message(self, message: str): self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
    def process_log_queue(self):
        try:
            while not self.log_queue.empty(): message = self.log_queue.get_nowait(); self.log_text.insert("end", message); self.log_text.see("end")
        finally:
            if not self.closing: self.window.after(100, self.process_log_queue)
    def clear_log(self): self.log_text.delete("1.0", "end"); self.log_message("Log cleared.")
    def save_log(self):
        try:
            with open("simulation_log.txt","w",encoding="utf-8") as f: f.write(self.log_text.get("1.0","end"))
            self.log_message("Log saved to simulation_log.txt")
        except Exception as e: self.log_message(f"Error saving log: {e}")
    def start_simulation_thread(self):
        if self.is_simulation_running: self.log_message("A simulation is already in progress."); return
        self.is_simulation_running=True; self.stop_simulation=False
        self.simulation_thread = threading.Thread(target=self.run_simulation, daemon=True); self.simulation_thread.start()
    def stop_simulation_thread(self):
        if self.is_simulation_running: self.stop_simulation=True; self.log_message("Stop signal sent to simulation."); self.sim_status_label.configure(text="Stopping simulation...")

if __name__ == "__main__":
    app = PatchAntennaDesigner()
    app.run()