# -*- coding: utf-8 -*-
import os
import tempfile
import time
from datetime import datetime
import math
import json
import traceback
import queue
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Tuple, List, Optional, Dict, Any

import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.mplot3d import Axes3D
import customtkinter as ctk

import ansys.aedt.core
from ansys.aedt.core import Desktop, Hfss

# ---------- Appearance ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class SimulationDialog(ctk.CTkToplevel):
    def __init__(self, parent, title="Simulation Progress"):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x200")
        self.transient(parent)
        self.grab_set()
        
        # Make dialog modal
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        
        self.label = ctk.CTkLabel(self, text="Starting simulation...", font=ctk.CTkFont(size=12))
        self.label.pack(pady=20)
        
        self.progress = ctk.CTkProgressBar(self, width=300)
        self.progress.pack(pady=10)
        self.progress.set(0)
        
        self.status_label = ctk.CTkLabel(self, text="0%", font=ctk.CTkFont(size=10))
        self.status_label.pack(pady=5)
        
        self.cancel_button = ctk.CTkButton(self, text="Cancel", command=self._on_cancel)
        self.cancel_button.pack(pady=10)
        
        self.cancelled = False
        
    def _on_cancel(self):
        self.cancelled = True
        self.destroy()
        
    def update_progress(self, value, message):
        self.progress.set(value)
        self.status_label.configure(text=f"{int(value*100)}% - {message}")
        self.update()


class ModernPatchAntennaDesigner(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.hfss = None
        self.desktop: Optional[Desktop] = None
        self.temp_folder = None
        self.project_path = ""
        self.project_display_name = "patch_array"
        self.design_base_name = "patch_array"
        self.log_queue = queue.Queue()
        self.is_simulation_running = False
        self.save_project = False
        self.stop_simulation = False
        self.simulation_data = None
        self.original_params = {}
        self.optimized = False
        self.optimization_history = []
        self.created_ports = []
        self.simulation_dialog = None
        self.progress_lock = threading.Lock()
        self.current_progress = 0.0
        self.current_status = "Starting..."

        # -------- User Parameters --------
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

        # -------- Calculated Parameters --------
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
        self.source_controls = {}
        
        # Now initialize Tkinter variables after the root window is created
        self.auto_refresh_var = ctk.BooleanVar(value=False)
        self.auto_refresh_job = None
        self.last_s11_analysis = None
        self.theta_cut = None
        self.phi_cut = None
        self.grid3d = None
        
        self.setup_gui()

    # ---------------- GUI ----------------
    def setup_gui(self):
        self.title("Patch Antenna Array Designer")
        self.geometry("1600x1000")

        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        header = ctk.CTkFrame(self, height=80, fg_color=("gray85", "gray20"))
        header.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header.grid_propagate(False)
        
        title_label = ctk.CTkLabel(
            header, 
            text="Patch Antenna Array Designer",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=("gray10", "gray90")
        )
        title_label.pack(pady=20)

        # Tab view
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Add tabs
        tabs = ["Design Parameters", "Simulation", "Results", "Log", "Beamforming"]
        for tab_name in tabs:
            self.tabview.add(tab_name)
            self.tabview.tab(tab_name).grid_columnconfigure(0, weight=1)

        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()
        self.setup_beamforming_tab()

        # Status bar
        status = ctk.CTkFrame(self, height=40)
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        
        self.status_label = ctk.CTkLabel(
            status, 
            text="Ready to calculate parameters",
            font=ctk.CTkFont(weight="bold")
        )
        self.status_label.pack(pady=8)

        self.process_log_queue()

    def create_section(self, parent, title, row, column, padx=10, pady=10):
        section = ctk.CTkFrame(parent, fg_color=("gray92", "gray18"))
        section.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
        section.grid_columnconfigure(0, weight=1)
        
        title_label = ctk.CTkLabel(
            section, 
            text=title,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=("gray20", "gray80")
        )
        title_label.grid(row=0, column=0, sticky="w", padx=15, pady=(10, 6))
        
        separator = ctk.CTkFrame(section, height=2, fg_color=("gray70", "gray30"))
        separator.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        
        return section

    def setup_parameters_tab(self):
        tab = self.tabview.tab("Design Parameters")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        # Create scrollable frame
        main_frame = ctk.CTkScrollableFrame(tab)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main_frame.grid_columnconfigure(0, weight=1)

        # Antenna Parameters section
        sec_ant = self.create_section(main_frame, "Antenna Parameters", 0, 0)
        self.entries = []
        row_idx = 2

        def add_entry(section, label, key, value, row, combo=None, check=False, tooltip=None):
            frame = ctk.CTkFrame(section, fg_color="transparent")
            frame.grid(row=row, column=0, sticky="ew", padx=15, pady=6)
            frame.grid_columnconfigure(0, weight=1)
            
            ctk.CTkLabel(
                frame, 
                text=label, 
                font=ctk.CTkFont(weight="bold"),
                width=220
            ).grid(row=0, column=0, sticky="w")
            
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(frame, values=combo, variable=var, width=220)
                widget.grid(row=0, column=1, sticky="w")
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(frame, text="", variable=var, width=30)
                widget.grid(row=0, column=1, sticky="w")
                self.entries.append((key, var))
            else:
                widget = ctk.CTkEntry(frame, width=220)
                widget.insert(0, str(value))
                widget.grid(row=0, column=1, sticky="w")
                self.entries.append((key, widget))
            
            # Add tooltip if provided
            if tooltip:
                self.create_tooltip(widget, tooltip)
            
            return row + 1

        # Add antenna parameters
        row_idx = add_entry(sec_ant, "Central Frequency (GHz):", "frequency", self.params["frequency"], row_idx,
                          tooltip="Operating frequency of the antenna array")
        row_idx = add_entry(sec_ant, "Desired Gain (dBi):", "gain", self.params["gain"], row_idx,
                          tooltip="Target gain for the antenna array")
        row_idx = add_entry(sec_ant, "Sweep Start (GHz):", "sweep_start", self.params["sweep_start"], row_idx,
                          tooltip="Start frequency for simulation sweep")
        row_idx = add_entry(sec_ant, "Sweep Stop (GHz):", "sweep_stop", self.params["sweep_stop"], row_idx,
                          tooltip="Stop frequency for simulation sweep")
        row_idx = add_entry(
            sec_ant, "Patch Spacing:", "spacing_type", self.params["spacing_type"], row_idx,
            combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"],
            tooltip="Spacing between patch elements as a fraction of wavelength"
        )

        # Substrate Parameters section
        sec_sub = self.create_section(main_frame, "Substrate Parameters", 1, 0)
        row_idx = 2
        
        row_idx = add_entry(
            sec_sub, "Substrate Material:", "substrate_material",
            self.params["substrate_material"], row_idx,
            combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"],
            tooltip="Material used for the substrate"
        )
        row_idx = add_entry(sec_sub, "Relative Permittivity (εr):", "er", self.params["er"], row_idx,
                          tooltip="Dielectric constant of the substrate material")
        row_idx = add_entry(sec_sub, "Loss Tangent (tan δ):", "tan_d", self.params["tan_d"], row_idx,
                          tooltip="Loss tangent of the substrate material")
        row_idx = add_entry(sec_sub, "Substrate Thickness (mm):", "substrate_thickness", self.params["substrate_thickness"], row_idx,
                          tooltip="Thickness of the substrate in millimeters")
        row_idx = add_entry(sec_sub, "Metal Thickness (mm):", "metal_thickness", self.params["metal_thickness"], row_idx,
                          tooltip="Thickness of the metal traces in millimeters")

        # Coaxial Feed Parameters section
        sec_coax = self.create_section(main_frame, "Coaxial Feed Parameters", 2, 0)
        row_idx = 2
        
        row_idx = add_entry(
            sec_coax, "Feed position type:", "feed_position", self.params["feed_position"], row_idx,
            combo=["inset", "edge"],
            tooltip="Type of feed configuration for the patch antenna"
        )
        row_idx = add_entry(sec_coax, "Feed relative X (0..1):", "feed_rel_x", self.params["feed_rel_x"], row_idx,
                          tooltip="Relative position of the feed along the patch width (0 to 1)")
        row_idx = add_entry(sec_coax, "Inner radius a (mm):", "probe_radius", self.params["probe_radius"], row_idx,
                          tooltip="Inner conductor radius of the coaxial feed")
        row_idx = add_entry(sec_coax, "b/a ratio:", "coax_ba_ratio", self.params["coax_ba_ratio"], row_idx,
                          tooltip="Ratio of outer to inner conductor radius")
        row_idx = add_entry(sec_coax, "Shield wall (mm):", "coax_wall_thickness", self.params["coax_wall_thickness"], row_idx,
                          tooltip="Thickness of the coaxial shield wall")
        row_idx = add_entry(sec_coax, "Port length below GND Lp (mm):", "coax_port_length", self.params["coax_port_length"], row_idx,
                          tooltip="Length of the coaxial port below ground")
        row_idx = add_entry(sec_coax, "Anti-pad clearance (mm):", "antipad_clearance", self.params["antipad_clearance"], row_idx,
                          tooltip="Clearance around the coaxial feed in the ground plane")

        # Simulation Settings section
        sec_sim = self.create_section(main_frame, "Simulation Settings", 3, 0)
        row_idx = 2
        
        row_idx = add_entry(sec_sim, "CPU Cores:", "cores", self.params["cores"], row_idx,
                          tooltip="Number of CPU cores to use for simulation")
        row_idx = add_entry(sec_sim, "Show HFSS Interface:", "show_gui", not self.params["non_graphical"], row_idx, check=True,
                          tooltip="Display HFSS graphical interface during simulation")
        row_idx = add_entry(sec_sim, "Save Project:", "save_project", self.save_project, row_idx, check=True,
                          tooltip="Save the HFSS project after simulation")
        row_idx = add_entry(
            sec_sim, "Sweep Type:", "sweep_type", self.params["sweep_type"], row_idx,
            combo=["Discrete", "Interpolating", "Fast"],
            tooltip="Type of frequency sweep for simulation"
        )
        row_idx = add_entry(sec_sim, "Discrete Step (GHz):", "sweep_step", self.params["sweep_step"], row_idx,
                          tooltip="Frequency step for discrete sweep (GHz)")
        row_idx = add_entry(sec_sim, "3D Theta step (deg):", "theta_step", self.params["theta_step"], row_idx,
                          tooltip="Step size for theta angle in 3D radiation patterns")
        row_idx = add_entry(sec_sim, "3D Phi step (deg):", "phi_step", self.params["phi_step"], row_idx,
                          tooltip="Step size for phi angle in 3D radiation patterns")

        # Calculated Parameters section
        sec_calc = self.create_section(main_frame, "Calculated Parameters", 4, 0)
        
        grid_frame = ctk.CTkFrame(sec_calc)
        grid_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        grid_frame.columnconfigure((0, 1), weight=1)

        self.patches_label = ctk.CTkLabel(grid_frame, text="Number of Patches: 4", font=ctk.CTkFont(weight="bold"))
        self.patches_label.grid(row=0, column=0, sticky="w", pady=4)
        
        self.rows_cols_label = ctk.CTkLabel(grid_frame, text="Configuration: 2 x 2", font=ctk.CTkFont(weight="bold"))
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=4)
        
        self.spacing_label = ctk.CTkLabel(grid_frame, text="Spacing: -- mm", font=ctk.CTkFont(weight="bold"))
        self.spacing_label.grid(row=1, column=0, sticky="w", pady=4)
        
        self.dimensions_label = ctk.CTkLabel(grid_frame, text="Patch Dimensions: -- x -- mm", font=ctk.CTkFont(weight="bold"))
        self.dimensions_label.grid(row=1, column=1, sticky="w", pady=4)
        
        self.lambda_label = ctk.CTkLabel(grid_frame, text="Guided Wavelength: -- mm", font=ctk
                                         

# ... (código anterior permanece igual)

    def setup_simulation_tab(self):
        tab = self.tabview.tab("Simulation")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            main, 
            text="Simulation Control", 
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=10)

        # Button row
        btn_row = ctk.CTkFrame(main)
        btn_row.pack(pady=14)
        
        self.run_button = ctk.CTkButton(
            btn_row, 
            text="Run Simulation", 
            command=self.start_simulation,
            fg_color="#2E8B57", 
            hover_color="#3CB371", 
            height=40, 
            width=160
        )
        self.run_button.pack(side="left", padx=8)
        
        self.stop_button = ctk.CTkButton(
            btn_row, 
            text="Stop Simulation", 
            command=self.stop_simulation,
            fg_color="#DC143C", 
            hover_color="#FF4500",
            state="disabled", 
            height=40, 
            width=160
        )
        self.stop_button.pack(side="left", padx=8)

        # Progress bar
        progress_frame = ctk.CTkFrame(main)
        progress_frame.pack(fill="x", padx=50, pady=8)
        
        ctk.CTkLabel(
            progress_frame, 
            text="Simulation Progress:", 
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w")
        
        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=18)
        self.progress_bar.pack(fill="x", pady=6)
        self.progress_bar.set(0)

        # Status label
        self.sim_status_label = ctk.CTkLabel(
            main, 
            text="Simulation not started",
            font=ctk.CTkFont(weight="bold")
        )
        self.sim_status_label.pack(pady=8)

        # Note
        note = ctk.CTkFrame(main, fg_color=("gray90", "gray15"))
        note.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(
            note, 
            text="Tip: With post vars (p_i / ph_i) you can retune beams without re-solving.",
            font=ctk.CTkFont(size=12, slant="italic"),
            text_color=("gray40", "gray60")
        ).pack(padx=10, pady=10)

    def start_simulation(self):
        """Start simulation with modal dialog"""
        if self.is_simulation_running:
            self.log_message("Simulation is already running")
            return
            
        self.stop_simulation_flag = False
        self.is_simulation_running = True
        
        # Create and show modal dialog
        self.simulation_dialog = SimulationDialog(self, "Simulation Progress")
        
        # Start simulation in a separate thread
        self.simulation_thread = threading.Thread(target=self.run_simulation, daemon=True)
        self.simulation_thread.start()
        
        # Start checking progress
        self.check_simulation_progress()

    def check_simulation_progress(self):
        """Check simulation progress and update dialog"""
        if self.simulation_dialog and self.simulation_dialog.winfo_exists():
            with self.progress_lock:
                progress = self.current_progress
                status = self.current_status
            
            self.simulation_dialog.update_progress(progress, status)
            
            if self.is_simulation_running:
                self.after(100, self.check_simulation_progress)
            else:
                self.simulation_dialog.destroy()
                self.simulation_dialog = None
        else:
            self.simulation_dialog = None

    def stop_simulation(self):
        """Stop the simulation"""
        self.stop_simulation_flag = True
        self.log_message("Simulation stop requested")

    def run_simulation(self):
        """Run the simulation - this runs in a separate thread"""
        try:
            self.log_message("Starting simulation")
            self.run_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            
            with self.progress_lock:
                self.current_progress = 0.0
                self.current_status = "Initializing..."
            
            # Get parameters
            if not self.get_parameters():
                self.log_message("Invalid parameters. Aborting.")
                return
                
            # Calculate parameters if not already done
            if self.calculated_params["num_patches"] < 1:
                self.calculate_parameters()

            # Open or create project
            self._open_or_create_project()
            
            with self.progress_lock:
                self.current_progress = 0.25
                self.current_status = "Project created"

            # Set model units
            self.hfss.modeler.model_units = "mm"
            self.log_message("Model units set to: mm")

            # Handle substrate material
            sub_name = self.params["substrate_material"]
            if not self.hfss.materials.checkifmaterialexists(sub_name):
                sub_name = "Custom_Substrate"
                self._ensure_material(sub_name, float(self.params["er"]), float(self.params["tan_d"]))

            # Get calculated parameters
            L = float(self.calculated_params["patch_length"])
            W = float(self.calculated_params["patch_width"])
            spacing = float(self.calculated_params["spacing"])
            rows = int(self.calculated_params["rows"])
            cols = int(self.calculated_params["cols"])
            h_sub = float(self.params["substrate_thickness"])
            sub_w = float(self.calculated_params["substrate_width"])
            sub_l = float(self.calculated_params["substrate_length"])

            # Set design variables
            self._set_design_variables(L, W, spacing, rows, cols, h_sub, sub_w, sub_l)

            # Create substrate
            self.log_message("Creating substrate")
            substrate = self.hfss.modeler.create_box(
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL", "h_sub"],
                name="Substrate",
                material=sub_name
            )
            
            # Create ground plane
            self.log_message("Creating ground plane")
            ground = self.hfss.modeler.create_rectangle(
                orientation="XY",
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL"],
                name="Ground",
                material="copper"
            )

            # Create patches
            self.log_message(f"Creating {rows*cols} patches in {rows}x{cols} configuration")
            patches = []
            total_w = cols * W + (cols - 1) * spacing
            total_l = rows * L + (rows - 1) * spacing
            start_x = -total_w / 2 + W / 2
            start_y = -total_l / 2 + L / 2
            
            with self.progress_lock:
                self.current_progress = 0.35
                self.current_status = "Creating patches..."

            # Create each patch
            count = 0
            for r in range(rows):
                for c in range(cols):
                    if self.stop_simulation_flag:
                        self.log_message("Simulation stopped by user")
                        return
                        
                    count += 1
                    patch_name = f"Patch_{count}"
                    cx = start_x + c * (W + spacing)
                    cy = start_y + r * (L + spacing)

                    origin = [cx - W / 2, cy - L / 2, "h_sub"]
                    self.log_message(f"Creating patch {count} at ({r}, {c})")

                    # Create patch
                    patch = self.hfss.modeler.create_rectangle(
                        orientation="XY",
                        origin=origin,
                        sizes=["patchW", "patchL"],
                        name=patch_name,
                        material="copper"
                    )
                    patches.append(patch)

                    # Calculate feed position
                    if self.params["feed_position"] == "edge":
                        y_feed = cy - 0.5 * L + 0.02 * L
                    else:
                        y_feed = cy - 0.5 * L + 0.30 * L
                        
                    relx = float(self.params["feed_rel_x"])
                    relx = min(max(relx, 0.0), 1.0)
                    x_feed = cx - 0.5 * W + relx * W

                    # Create feed pad
                    pad = self.hfss.modeler.create_circle(
                        orientation="XY",
                        origin=[x_feed, y_feed, "h_sub"],
                        radius="a",
                        name=f"{patch_name}_Pad",
                        material="copper"
                    )
                    
                    # Unite patch and pad
                    try:
                        self.hfss.modeler.unite([patch, pad])
                    except Exception:
                        pass

                    # Create coaxial feed
                    self._create_coax_feed_lumped(
                        ground=ground,
                        substrate=substrate,
                        x_feed=x_feed,
                        y_feed=y_feed,
                        name_prefix=f"P{count}"
                    )
                    
                    # Update progress
                    progress = 0.35 + 0.25 * (count / float(rows * cols))
                    with self.progress_lock:
                        self.current_progress = progress
                        self.current_status = f"Creating patch {count}/{rows*cols}"

            # Check if simulation was stopped
            if self.stop_simulation_flag:
                self.log_message("Simulation stopped by user")
                return

            # Assign PerfectE boundary to ground and patches
            try:
                names = [ground.name] + [p.name for p in patches]
                self.hfss.assign_perfecte_to_sheets(names)
                self.log_message(f"PerfectE assigned to: {names}")
            except Exception as e:
                self.log_message(f"PerfectE assignment warning: {e}")

            # Create air region and radiation boundary
            self.log_message("Creating air region + radiation boundary")
            lambda0_mm = self.c / (self.params["sweep_start"] * 1e9) * 1000.0
            pad_mm = float(lambda0_mm) / 4.0
            
            region = self.hfss.modeler.create_region(
                [pad_mm, pad_mm, pad_mm, pad_mm, pad_mm, pad_mm], 
                is_percentage=False
            )
            
            self.hfss.assign_radiation_boundary_to_objects(region)
            
            with self.progress_lock:
                self.current_progress = 0.65
                self.current_status = "Creating simulation setup..."

            # Create simulation setup
            self.log_message("Creating simulation setup")
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02

            # Create frequency sweep
            self.log_message(f"Creating frequency sweep: {self.params['sweep_type']}")
            stype = self.params["sweep_type"]
            
            try:
                # Delete existing sweep if any
                try:
                    sw = setup.get_sweep("Sweep1")
                    if sw:
                        sw.delete()
                except Exception:
                    pass

                # Create appropriate sweep type
                if stype == "Discrete":
                    step = float(self.params["sweep_step"])
                    setup.create_linear_step_sweep(
                        unit="GHz",
                        start_frequency=self.params["sweep_start"],
                        stop_frequency=self.params["sweep_stop"],
                        step_size=step,
                        name="Sweep1"
                    )
                elif stype == "Fast":
                    setup.create_frequency_sweep(
                        unit="GHz",
                        name="Sweep1",
                        start_frequency=self.params["sweep_start"],
                        stop_frequency=self.params["sweep_stop"],
                        sweep_type="Fast"
                    )
                else:  # Interpolating
                    setup.create_frequency_sweep(
                        unit="GHz",
                        name="Sweep1",
                        start_frequency=self.params["sweep_start"],
                        stop_frequency=self.params["sweep_stop"],
                        sweep_type="Interpolating"
                    )
                    
            except Exception as e:
                self.log_message(f"Sweep creation warning: {e}")

            # Check excitations
            excitations = self._list_excitations()
            self.log_message(f"Excitations created: {len(excitations)} -> {excitations}")
            
            if not excitations:
                self.sim_status_label.configure(text="No excitations defined")
                self.log_message("No excitations found. Aborting before solve.")
                return

            # Validate design
            self.log_message("Validating design")
            try:
                _ = self.hfss.validate_full_design()
            except Exception as e:
                self.log_message(f"Validation warning: {e}")

            # Start analysis
            self.log_message("Starting analysis")
            if self.save_project:
                self.hfss.save_project()
                
            with self.progress_lock:
                self.current_progress = 0.7
                self.current_status = "Running simulation..."
                
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])

            # Check if simulation was stopped
            if self.stop_simulation_flag:
                self.log_message("Simulation stopped by user")
                return

            # Perform post-processing
            self._postprocess_after_solve()

            # Process results
            with self.progress_lock:
                self.current_progress = 0.9
                self.current_status = "Processing results..."
                
            self.log_message("Processing results")
            self.plot_results()
            
            # Finalize
            with self.progress_lock:
                self.current_progress = 1.0
                self.current_status = "Simulation completed"
                
            self.log_message("Simulation completed successfully")
            
        except Exception as e:
            self.log_message(f"Error in simulation: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            with self.progress_lock:
                self.current_status = f"Simulation error: {e}"
            
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    # ... (rest of the methods remain mostly the same, but ensure they use the new progress system)

    def cleanup(self):
        """Clean up resources."""
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
                    self.log_message(f"Error cleaning up temporary files: {e}")
                    
        except Exception as e:
            self.log_message(f"Error during cleanup: {e}")

    def on_closing(self):
        """Handle window closing event."""
        self.log_message("Application closing...")
        self.cleanup()
        self.quit()
        self.destroy()

if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.mainloop()