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

# Configuração do tema
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ModernPatchAntennaDesigner:
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
            "substrate_thickness": 0.5, "metal_thickness": 0.035,
            "er": 3.55, "tan_d": 0.0027, "feed_position": "inset",
            "probe_radius": 0.4, "coax_er": 2.1, "coax_wall_thickness": 0.2,
            "coax_port_length": 3.0, "antipad_clearance": 0.2
        }

        self.calculated_params: Dict[str, Any] = {
            "num_patches": 4, "spacing": 21.0, "patch_length": 9.04, "patch_width": 11.5,
            "rows": 2, "cols": 2, "lambda_g": 15.9, "feed_offset": 3.0,
            "substrate_width": 75.0, "substrate_length": 70.0
        }
        
        self.simulation_data: Dict[str, Any] = {}
        self.c: float = 299792458.0
        self.setup_gui()

    # --- Configuração da Interface Gráfica ---
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        self.window.geometry("1400x950")
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.window.grid_columnconfigure(0, weight=1); self.window.grid_rowconfigure(1, weight=1)
        
        header_frame = ctk.CTkFrame(self.window, height=80, fg_color=("gray85", "gray20"))
        header_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header_frame.pack_propagate(False)
        ctk.CTkLabel(header_frame, text="Patch Antenna Array Designer", font=ctk.CTkFont(size=24, weight="bold"), text_color=("gray10", "gray90")).pack(pady=20)
        
        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        tabs = ["Design Parameters", "Simulation", "Results", "Log"]
        for tab_name in tabs: self.tabview.add(tab_name); self.tabview.tab(tab_name).grid_columnconfigure(0, weight=1)
        
        self.setup_parameters_tab(); self.setup_simulation_tab(); self.setup_results_tab(); self.setup_log_tab()
        
        status_frame = ctk.CTkFrame(self.window, height=40)
        status_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 5))
        status_frame.pack_propagate(False)
        self.status_label = ctk.CTkLabel(status_frame, text="Ready to calculate parameters", font=ctk.CTkFont(weight="bold"))
        self.status_label.pack(pady=10)
        
        self.process_log_queue()

    def create_param_entry(self, parent, key, label, row, combo_values=None, is_bool=False):
        ctk.CTkLabel(parent, text=label, font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, padx=15, pady=5, sticky="w")
        if combo_values:
            widget = ctk.CTkComboBox(parent, values=combo_values, width=200)
            widget.set(self.params[key])
        elif is_bool:
            widget = ctk.CTkCheckBox(parent, text="", onvalue=True, offvalue=False)
            current_value = self.params.get(key, False)
            if current_value: widget.select() 
            else: widget.deselect()
        else:
            widget = ctk.CTkEntry(parent, width=200)
            widget.insert(0, str(self.params[key]))
        widget.grid(row=row, column=1, padx=15, pady=5, sticky="w")
        self.entries[key] = widget

    def setup_parameters_tab(self):
        self.entries: Dict[str, Any] = {}
        tab = self.tabview.tab("Design Parameters")
        tab.grid_columnconfigure(0, weight=1); tab.grid_rowconfigure(0, weight=1)
        main_frame = ctk.CTkScrollableFrame(tab); main_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5); main_frame.grid_columnconfigure(0, weight=1)

        sections = {
            "Antenna": [("frequency", "Central Frequency (GHz)"), ("gain", "Target Gain (dBi)"), ("sweep_start", "Sweep Start (GHz)"), ("sweep_stop", "Sweep Stop (GHz)"), ("spacing_type", "Patch Spacing", ["0.6*lambda", "0.7*lambda", "0.8*lambda", "lambda/2", "lambda"])],
            "Substrate": [("substrate_material", "Material", ["Rogers RO4003C (tm)", "FR4_epoxy", "Duroid 5880 (tm)", "Air"]), ("er", "Relative Permittivity (εr)"), ("tan_d", "Loss Tangent (tan δ)"), ("substrate_thickness", "Substrate Thickness (mm)"), ("metal_thickness", "Metal Thickness (mm)")],
            "Coaxial Feed": [("feed_position", "Feed Position Type", ["inset", "edge"]), ("probe_radius", "Probe Radius 'a' (mm)"), ("coax_er", "Coax εr (e.g., PTFE)"), ("coax_wall_thickness", "Shield Wall Thickness (mm)"), ("coax_port_length", "Port Length below GND (mm)"), ("antipad_clearance", "Anti-pad Clearance (mm)")],
            "Simulation": [("cores", "CPU Cores"), ("non_graphical", "Run Non-Graphical", None, True), ("save_project", "Save Project on Exit", None, True)]
        }
        
        current_row = 0
        for section_title, params_list in sections.items():
            section_frame = ctk.CTkFrame(main_frame, fg_color=("gray92", "gray18"))
            section_frame.grid(row=current_row, column=0, sticky="ew", padx=10, pady=10); section_frame.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(section_frame, text=section_title, font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=15, pady=(10, 5))
            ctk.CTkFrame(section_frame, height=2, fg_color=("gray70", "gray30")).grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 5)) # Separador
            row_in_section = 2
            for param_info in params_list:
                key, label, *rest = param_info
                combo = rest[0] if rest and isinstance(rest[0], list) else None
                is_bool = rest[1] if len(rest) > 1 and isinstance(rest[1], bool) else False
                self.create_param_entry(section_frame, key, label, row_in_section, combo, is_bool); row_in_section += 1
            current_row += 1

        calc_section = ctk.CTkFrame(main_frame, fg_color=("gray92", "gray18")); calc_section.grid(row=current_row, column=0, sticky="ew", padx=10, pady=10); calc_section.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(calc_section, text="Calculated Parameters", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, sticky="w", padx=15, pady=(10,5))
        ctk.CTkFrame(calc_section, height=2, fg_color=("gray70", "gray30")).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        calc_grid = ctk.CTkFrame(calc_section, fg_color="transparent"); calc_grid.grid(row=2, padx=15, pady=10, sticky="ew"); calc_grid.columnconfigure((0, 1), weight=1)
        self.calc_labels = { "patches": ctk.CTkLabel(calc_grid, font=ctk.CTkFont(weight="bold")), "rows_cols": ctk.CTkLabel(calc_grid, font=ctk.CTkFont(weight="bold")), "spacing": ctk.CTkLabel(calc_grid, font=ctk.CTkFont(weight="bold")), "dimensions": ctk.CTkLabel(calc_grid, font=ctk.CTkFont(weight="bold")), "lambda_g": ctk.CTkLabel(calc_grid, font=ctk.CTkFont(weight="bold")), "feed_offset": ctk.CTkLabel(calc_grid, font=ctk.CTkFont(weight="bold")), "substrate_dims": ctk.CTkLabel(calc_grid, font=ctk.CTkFont(weight="bold")) }
        for i, label in enumerate(self.calc_labels.values()): label.grid(row=i//2, column=i%2, sticky="w", pady=2, padx=5)
        self.calc_labels["substrate_dims"].grid(row=len(self.calc_labels)//2, column=0, columnspan=2)
        
        button_frame = ctk.CTkFrame(main_frame); button_frame.grid(row=current_row + 1, pady=15)
        ctk.CTkButton(button_frame, text="Calculate Parameters", command=self.calculate_parameters, fg_color="#2E8B57", hover_color="#3CB371").pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Save", command=self.save_parameters).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Load", command=self.load_parameters).pack(side="left", padx=10)
        self.update_calculated_params_display()
    
    def setup_simulation_tab(self):
        tab=self.tabview.tab("Simulation");tab.grid_columnconfigure(0,weight=1);tab.grid_rowconfigure(0,weight=1)
        mf=ctk.CTkFrame(tab);mf.grid(row=0,column=0,sticky="nsew",padx=10,pady=10);mf.grid_columnconfigure(0,weight=1)
        ctk.CTkLabel(mf,text="Simulation Control",font=ctk.CTkFont(size=16,weight="bold")).pack(pady=10)
        bf=ctk.CTkFrame(mf);bf.pack(pady=20)
        self.run_button=ctk.CTkButton(bf,text="Run Simulation",command=self.start_simulation_thread,fg_color="#2E8B57",hover_color="#3CB371",height=40,width=150);self.run_button.pack(side="left",padx=10)
        self.stop_button=ctk.CTkButton(bf,text="Stop Simulation",command=self.stop_simulation_thread,fg_color="#DC143C",hover_color="#FF4500",height=40,width=150,state="disabled");self.stop_button.pack(side="left",padx=10)
        pf=ctk.CTkFrame(mf);pf.pack(fill="x",padx=50,pady=10)
        ctk.CTkLabel(pf,text="Simulation Progress:",font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.progress_bar=ctk.CTkProgressBar(pf,height=20);self.progress_bar.pack(fill="x",pady=5);self.progress_bar.set(0)
        self.sim_status_label=ctk.CTkLabel(mf,text="Simulation not started",font=ctk.CTkFont(weight="bold"));self.sim_status_label.pack(pady=10)

    def setup_results_tab(self):
        tab=self.tabview.tab("Results");tab.grid_columnconfigure(0,weight=1);tab.grid_rowconfigure(1,weight=1)
        mf=ctk.CTkFrame(tab);mf.grid(row=0,column=0,sticky="nsew",padx=10,pady=10);mf.grid_columnconfigure(0,weight=1);mf.grid_rowconfigure(1,weight=1)
        top_frame = ctk.CTkFrame(mf); top_frame.grid(row=0, column=0, sticky="ew", pady=(10,0), padx=10)
        ctk.CTkLabel(top_frame, text="S-Parameter:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(10,5))
        self.sparam_selector = ctk.CTkOptionMenu(top_frame, values=["N/A"], command=self.plot_selected_sparam); self.sparam_selector.pack(side="left", padx=5)
        gf=ctk.CTkFrame(mf);gf.grid(row=1,column=0,sticky="nsew",padx=10,pady=10)
        self.fig, self.ax = plt.subplots(figsize=(8, 6)); self.update_plot_theme()
        self.canvas=FigureCanvasTkAgg(self.fig,master=gf);self.canvas.get_tk_widget().pack(fill="both",expand=True)
        ef=ctk.CTkFrame(mf);ef.grid(row=2,column=0,pady=10)
        ctk.CTkButton(ef,text="Export CSV",command=self.export_csv).pack(side="left",padx=10)
        ctk.CTkButton(ef,text="Export PNG",command=self.export_png).pack(side="left",padx=10)

    def setup_log_tab(self):
        tab = self.tabview.tab("Log");tab.grid_columnconfigure(0, weight=1);tab.grid_rowconfigure(1, weight=1)
        mf = ctk.CTkFrame(tab);mf.grid(row=0, column=0, sticky="nsew", padx=10, pady=10);mf.grid_columnconfigure(0, weight=1);mf.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(mf, text="Simulation Log", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, pady=10)
        self.log_text = ctk.CTkTextbox(mf, width=900, height=500, font=ctk.CTkFont(family="Consolas"));self.log_text.grid(row=1, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", f"Log started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        bf = ctk.CTkFrame(mf);bf.grid(row=2, pady=10)
        ctk.CTkButton(bf, text="Clear Log", command=self.clear_log).pack(side="left", padx=10)
        ctk.CTkButton(bf, text="Save Log", command=self.save_log).pack(side="left", padx=10)
    
    # --- Funções de Lógica e Cálculo ---
    def get_parameters(self) -> bool:
        self.log_message("Getting parameters from UI...")
        try:
            for key, widget in self.entries.items():
                value = widget.get()
                if isinstance(widget, ctk.CTkCheckBox):
                    if key == 'non_graphical': self.params[key] = not value
                    else: self.save_project = value
                elif isinstance(widget, ctk.CTkComboBox): self.params[key] = value
                else:
                    if key == "cores": self.params[key] = int(value)
                    else: self.params[key] = float(value)
            self.log_message("Parameters retrieved successfully.")
            return True
        except (ValueError, TypeError) as e:
            msg = f"Invalid input for '{key}': {e}"; self.log_message(f"ERROR: {msg}"); self.status_label.configure(text=msg)
            return False

    def calculate_patch_dimensions(self) -> Tuple[float, float, float]:
        f, er, h = self.params["frequency"] * 1e9, self.params["er"], self.params["substrate_thickness"] / 1000.0
        W = self.c / (2 * f) * math.sqrt(2 / (er + 1))
        eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W)**(-0.5)
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        L = self.c / (2 * f * math.sqrt(eeff)) - 2 * dL
        lambda_g = (self.c / (f * math.sqrt(eeff))) * 1000.0
        return L * 1000.0, W * 1000.0, lambda_g
    
    def _calculate_array_dims(self, n_min: int) -> Tuple[int, int]:
        if n_min <= 1: return 1, 1
        sqrt_n = math.isqrt(n_min)
        rows = sqrt_n
        while rows > 1:
            if n_min % rows == 0: break
            rows -= 1
        cols = n_min // rows if rows > 0 else n_min
        return max(rows, cols), min(rows, cols)

    def calculate_parameters(self):
        if not self.get_parameters(): return
        self.log_message("Calculating design parameters...")
        try:
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions()
            self.calculated_params.update({"patch_length": L_mm, "patch_width": W_mm, "lambda_g": lambda_g_mm})
            L_m = L_mm / 1000.0
            try: y0 = (L_m / math.pi) * math.acos(math.sqrt(50.0 / 240.0)); self.calculated_params["feed_offset"] = y0 * 1000.0
            except ValueError: self.calculated_params["feed_offset"] = L_mm / 3.0
            gain_single, target_gain = 6.5, self.params["gain"]
            n_min = max(1, math.ceil(10**((target_gain - gain_single) / 10.0)))
            rows, cols = self._calculate_array_dims(n_min)
            self.calculated_params.update({"num_patches": rows * cols, "rows": rows, "cols": cols})
            spacing_str = self.params["spacing_type"].replace("lambda", "1.0")
            spacing_factor = eval(spacing_str) if "/" in spacing_str else float(spacing_str.split('*')[0])
            lambda0_mm = (self.c / (self.params["frequency"] * 1e9)) * 1000.0
            self.calculated_params["spacing"] = spacing_factor * lambda0_mm
            self.calculate_substrate_size()
            self.update_calculated_params_display()
            self.status_label.configure(text="Parameters calculated. Ready for simulation.")
            self.log_message("Parameter calculation finished.")
        except Exception as e:
            msg = f"Calculation Error: {e}"; self.log_message(f"ERROR: {msg}\n{traceback.format_exc()}"); self.status_label.configure(text=msg)
            
    def calculate_substrate_size(self):
        cp = self.calculated_params
        total_w = cp["cols"] * cp["patch_width"] + (cp["cols"] - 1) * cp["spacing"]
        total_l = cp["rows"] * cp["patch_length"] + (cp["rows"] - 1) * cp["spacing"]
        margin = max(total_w, total_l) * 0.2
        self.calculated_params["substrate_width"] = total_w + 2 * margin
        self.calculated_params["substrate_length"] = total_l + 2 * margin
    
    def update_calculated_params_display(self):
        cp = self.calculated_params
        self.calc_labels["patches"].configure(text=f"Patches: {cp.get('num_patches', 'N/A')}")
        self.calc_labels["rows_cols"].configure(text=f"Grid: {cp.get('rows', 'N/A')} x {cp.get('cols', 'N/A')}")
        self.calc_labels["spacing"].configure(text=f"Spacing: {cp.get('spacing', 0):.2f} mm")
        self.calc_labels["dimensions"].configure(text=f"Patch (LxW): {cp.get('patch_length', 0):.2f} x {cp.get('patch_width', 0):.2f} mm")
        self.calc_labels["lambda_g"].configure(text=f"λg: {cp.get('lambda_g', 0):.2f} mm")
        self.calc_labels["feed_offset"].configure(text=f"Feed Offset: {cp.get('feed_offset', 0):.2f} mm")
        self.calc_labels["substrate_dims"].configure(text=f"Substrate (LxW): {cp.get('substrate_length', 0):.2f} x {cp.get('substrate_width', 0):.2f} mm")
    
    # --- Funções de Simulação com PyAEDT ---
    def run_simulation(self):
        try:
            self.log_message("Starting simulation thread...")
            self.run_button.configure(state="disabled"); self.stop_button.configure(state="normal")
            self.sim_status_label.configure(text="Initializing AEDT..."); self.progress_bar.set(0)
            
            with tempfile.TemporaryDirectory(suffix=".ansys") as self.temp_folder:
                self.project_name = os.path.join(self.temp_folder, "patch_array.aedt")
                with ansys.aedt.core.Desktop(self.params["aedt_version"], self.params['non_graphical'], new_desktop=True) as self.desktop:
                    self.hfss = ansys.aedt.core.Hfss(project=self.project_name, design="patch_array", solution_type="DrivenModal")
                    self.log_message(f"AEDT {self.params['aedt_version']} and project initialized.")
                    self.hfss.modeler.model_units = "mm"
                    
                    self.sim_status_label.configure(text="Creating 3D Model..."); self.progress_bar.set(0.1)
                    self.create_geometry()
                    if self.stop_simulation: return

                    self.sim_status_label.configure(text="Configuring Analysis..."); self.progress_bar.set(0.6)
                    self.hfss.modeler.create_air_region(x_pos="padAir", y_pos="padAir", z_pos="padAir", x_neg="padAir", y_neg="padAir", z_neg="padAir")
                    setup = self.hfss.create_setup(name="Setup1")
                    setup.props["Frequency"] = "f0"; setup.props["MaximumPasses"] = 10
                    sweep = setup.create_linear_count_sweep(unit="GHz", start_frequency=self.params["sweep_start"], stop_frequency=self.params["sweep_stop"], num_of_freq_points=201)
                    
                    self.sim_status_label.configure(text="Running Simulation..."); self.progress_bar.set(0.7)
                    self.hfss.analyze_setup("Setup1", cores=self.params["cores"])
                    if self.stop_simulation: return

                    self.sim_status_label.configure(text="Processing Results..."); self.progress_bar.set(0.9)
                    self.update_results()
                    self.sim_status_label.configure(text="Simulation Completed."); self.progress_bar.set(1.0)
        except Exception as e:
            msg = f"Simulation failed: {e}"; self.log_message(f"FATAL ERROR: {msg}\n{traceback.format_exc()}"); self.sim_status_label.configure(text=msg)
        finally:
            self.cleanup()
            self.run_button.configure(state="normal"); self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    def _ensure_material(self, name, er, tan_d):
        if name not in self.hfss.materials.material_keys:
            self.log_message(f"Material '{name}' not found. Creating it...")
            new_mat = self.hfss.materials.add_material(name)
            new_mat.permittivity = er
            new_mat.dielectric_loss_tangent = tan_d
            new_mat.update()
            
    def _create_coax_feed(self, name_prefix, x_pos, y_pos):
        self.hfss.modeler.create_cylinder("Z", [x_pos, y_pos, "-Lp"], "a", "h_sub+Lp+eps", name=f"{name_prefix}_Pin", material="copper")
        ptfe = self.hfss.modeler.create_cylinder("Z", [x_pos, y_pos, "-Lp"], "b", "Lp", name=f"{name_prefix}_PTFE_solid", material="PTFE_Custom")
        self.hfss.modeler.subtract(ptfe, f"{name_prefix}_Pin", keep_originals=False)
        shield = self.hfss.modeler.create_cylinder("Z", [x_pos, y_pos, "-Lp"], "b+wall", "Lp", name=f"{name_prefix}_Shield_solid", material="copper")
        shield_void = self.hfss.modeler.create_cylinder("Z", [x_pos, y_pos, "-Lp"], "b", "Lp", name=f"{name_prefix}_Shield_void")
        self.hfss.modeler.subtract(shield, shield_void, keep_originals=False)
        sub_hole = self.hfss.modeler.create_cylinder("Z", [x_pos, y_pos, 0], "b+clear", "h_sub", name=f"{name_prefix}_SubHole")
        self.hfss.modeler.subtract("Substrate", sub_hole, keep_originals=False)
        gnd_hole = self.hfss.modeler.create_circle("XY", [x_pos, y_pos, 0], "b+clear", name=f"{name_prefix}_GndHole")
        self.hfss.modeler.subtract("Ground", gnd_hole, keep_originals=False)
        port_cap = self.hfss.modeler.create_circle("XY", [x_pos, y_pos, "-Lp"], "b", name=f"Port_{name_prefix}")
        self.hfss.wave_port(port_cap.name, impedance=50, name=name_prefix)

    def create_geometry(self):
        self._ensure_material(self.params["substrate_material"], self.params["er"], self.params["tan_d"])
        self._ensure_material("PTFE_Custom", self.params["coax_er"], 0.0002)
        
        b_val = self.params['probe_radius'] * math.exp(50.0 * math.sqrt(self.params['coax_er']) / 60.0)
        lambda0_mm = self.c / (self.params["frequency"] * 1e9) * 1000.0
        
        variables = {
            "f0": f"{self.params['frequency']}GHz", "h_sub": f"{self.params['substrate_thickness']}mm", "patchL": f"{self.calculated_params['patch_length']}mm",
            "patchW": f"{self.calculated_params['patch_width']}mm", "spacing": f"{self.calculated_params['spacing']}mm", "rows": str(self.calculated_params['rows']),
            "cols": str(self.calculated_params['cols']), "subW": f"{self.calculated_params['substrate_width']}mm", "subL": f"{self.calculated_params['substrate_length']}mm",
            "a": f"{self.params['probe_radius']}mm", "b": f"{b_val}mm", "wall": f"{self.params['coax_wall_thickness']}mm", "Lp": f"{self.params['coax_port_length']}mm",
            "clear": f"{self.params['antipad_clearance']}mm", "eps": "0.001mm", "probeOfsY": f"{self.calculated_params['feed_offset']}mm", "padAir": f"{lambda0_mm / 2}mm"
        }
        for name, value in variables.items(): self.hfss[name] = value

        self.hfss.modeler.create_box(["-subW/2", "-subL/2", 0], ["subW", "subL", "h_sub"], name="Substrate", material=self.params["substrate_material"])
        gnd = self.hfss.modeler.create_rectangle("XY", ["-subW/2", "-subL/2", 0], ["subW", "subL"], name="Ground", material="copper")
        self.hfss.assign_perfecte_to_sheets(gnd)
        
        total_w_expr, total_l_expr = "cols*patchW + (cols-1)*spacing", "rows*patchL + (rows-1)*spacing"
        start_x_expr, start_y_expr = f"-({total_w_expr})/2 + patchW/2", f"-({total_l_expr})/2 + patchL/2"
        
        for r in range(self.calculated_params['rows']):
            for c in range(self.calculated_params['cols']):
                port_num = r * self.calculated_params['cols'] + c + 1
                port_name, patch_name = f"P{port_num}", f"Patch_{port_num}"
                cx, cy = f"{start_x_expr} + {c}*(patchW+spacing)", f"{start_y_expr} + {r}*(patchL+spacing)"
                patch = self.hfss.modeler.create_rectangle("XY", [f"{cx}-patchW/2", f"{cy}-patchL/2", "h_sub"], ["patchW", "patchL"], name=patch_name, material="copper")
                y_feed_expr = f"{cy} - patchL/2 + probeOfsY" if self.params['feed_position'] == 'inset' else f"{cy} - patchL/2"
                self.hfss.modeler.create_circle("XY", [cx, y_feed_expr, "h_sub"], "a", name=f"{patch_name}_Pad", material="copper")
                self.hfss.modeler.unite([patch_name, f"{patch_name}_Pad"])
                self._create_coax_feed(port_name, cx, y_feed_expr)

    # --- Funções de Pós-Processamento e Utilitários ---
    def update_results(self):
        self.log_message("Fetching S-Parameter data...")
        try:
            traces = self.hfss.post.get_solution_data_per_variation("S-Parameters", "Setup1 : Sweep1")
            if traces:
                self.simulation_data["s_params"] = traces
                param_list = sorted(traces.expressions)
                self.sparam_selector.configure(values=param_list)
                if param_list:
                    default_param = next((p for p in param_list if p.count(',') == 1 and p.split(',')[0] == p.split(',')[1]), param_list[0])
                    self.sparam_selector.set(default_param)
                    self.plot_selected_sparam(default_param)
        except Exception as e: self.log_message(f"Could not retrieve S-Parameter data: {e}")

    def plot_selected_sparam(self, choice):
        if "s_params" in self.simulation_data and self.simulation_data["s_params"]:
            self.ax.clear(); self.update_plot_theme()
            trace = self.simulation_data["s_params"]
            freqs, data = trace.primary_sweep_values, trace.data_db(choice)
            self.ax.plot(freqs, data, label=choice, color="#1f77b4")
            if choice.count(',') == 1 and choice.split(',')[0] == choice.split(',')[1]:
                self.ax.axhline(y=-10, color='red', linestyle='--', alpha=0.7, label='-10 dB Ref.')
            self.ax.set_xlabel("Frequency (GHz)"); self.ax.set_ylabel("Magnitude (dB)"); self.ax.set_title(f"{choice} Parameter")
            self.ax.legend(); self.ax.grid(True, alpha=0.3)
            self.canvas.draw()
            
    def update_plot_theme(self):
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg_color, text_color = ('#2B2B2B', 'white') if is_dark else ('#F0F0F0', 'black')
        self.fig.set_facecolor(bg_color); self.ax.set_facecolor(bg_color)
        self.ax.tick_params(colors=text_color, which='both')
        for spine in self.ax.spines.values(): spine.set_edgecolor(text_color)
        for item in [self.ax.xaxis.label, self.ax.yaxis.label, self.ax.title]: item.set_color(text_color)

    def on_closing(self):
        self.log_message("Application closing..."); self.closing = True
        self.cleanup(); self.window.quit(); self.window.destroy()

    def cleanup(self):
        if self.is_simulation_running: self.stop_simulation = True
        if self.hfss:
            try:
                if self.save_project: self.hfss.save_project()
                else: self.hfss.close_project(save=False)
                self.log_message("HFSS project closed.")
            except Exception as e: self.log_message(f"Error during project cleanup: {e}")

    def save_parameters(self):
        try: self.get_parameters(); f=open("antenna_design.json","w"); json.dump({"input":self.params,"calculated":self.calculated_params,"save_project":self.save_project},f,indent=4); f.close(); self.log_message("Design parameters saved.")
        except Exception as e: self.log_message(f"Error saving parameters: {e}")

    def load_parameters(self):
        try:
            with open("antenna_design.json", "r") as f: data = json.load(f)
            self.params = data.get("input", self.params); self.calculated_params = data.get("calculated", self.calculated_params); self.save_project = data.get("save_project", self.save_project)
            self.update_interface_from_params(); self.log_message("Parameters loaded.")
        except FileNotFoundError: self.log_message("Error: antenna_design.json not found.")
        except Exception as e: self.log_message(f"Error loading parameters: {e}")

    def update_interface_from_params(self):
        for key, widget in self.entries.items():
            if key in self.params:
                value = not self.params[key] if key == 'non_graphical' else self.params.get(key)
                if isinstance(widget, ctk.CTkCheckBox):
                    if value: widget.select()
                    else: widget.deselect()
                elif isinstance(widget, ctk.CTkComboBox): widget.set(str(value))
                else: widget.delete(0, "end"); widget.insert(0, str(value))
        if self.entries.get("save_project"):
            if self.save_project: self.entries["save_project"].select()
            else: self.entries["save_project"].deselect()
        self.update_calculated_params_display()
        
    def run(self): self.window.mainloop()
    def clear_log(self): self.log_text.delete("1.0", "end"); self.log_message("Log cleared.")
    def save_log(self):
        try:
            with open("simulation_log.txt","w",encoding="utf-8") as f: f.write(self.log_text.get("1.0","end"))
            self.log_message("Log saved to simulation_log.txt")
        except Exception as e: self.log_message(f"Error saving log: {e}")
    def log_message(self, message: str): self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
    def process_log_queue(self):
        try:
            while not self.log_queue.empty(): message = self.log_queue.get_nowait(); self.log_text.insert("end", message); self.log_text.see("end")
        finally:
            if not self.closing: self.window.after(100, self.process_log_queue)
    def start_simulation_thread(self):
        if self.is_simulation_running: self.log_message("A simulation is already in progress."); return
        self.is_simulation_running=True; self.stop_simulation=False
        self.simulation_thread = threading.Thread(target=self.run_simulation, daemon=True); self.simulation_thread.start()
    def stop_simulation_thread(self):
        if self.is_simulation_running: self.stop_simulation=True; self.log_message("Stop signal sent to simulation."); self.sim_status_label.configure(text="Stopping simulation...")
    def export_csv(self):
        if "s_params" in self.simulation_data:
            try:
                selected_sparam = self.sparam_selector.get()
                trace = self.simulation_data["s_params"]
                data_to_save = np.column_stack((trace.primary_sweep_values, trace.data_db(selected_sparam)))
                filename = f"results_{selected_sparam.replace('(','').replace(')','').replace(',','_')}.csv"
                np.savetxt(filename, data_to_save, delimiter=",", header=f"Frequency_GHz,{selected_sparam}_dB", comments="")
                self.log_message(f"Data exported to {filename}")
            except Exception as e: self.log_message(f"Error exporting CSV: {e}")
        else: self.log_message("No simulation data to export.")
    def export_png(self):
        try:
            selected_sparam = self.sparam_selector.get()
            filename = f"plot_{selected_sparam.replace('(','').replace(')','').replace(',','_')}.png"
            self.fig.savefig(filename, dpi=300, bbox_inches='tight', facecolor=self.fig.get_facecolor())
            self.log_message(f"Plot saved to {filename}")
        except Exception as e: self.log_message(f"Error saving plot: {e}")

if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()