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


class ModernPatchAntennaDesigner(ctk.CTk):
    def __init__(self):
        super().__init__()  # Initialize the root window first
        
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
        
        self.lambda_label = ctk.CTkLabel(grid_frame, text="Guided Wavelength: -- mm", font=ctk.CTkFont(weight="bold"))
        self.lambda_label.grid(row=2, column=0, sticky="w", pady=4)
        
        self.feed_offset_label = ctk.CTkLabel(grid_frame, text="Feed Offset (y): -- mm", font=ctk.CTkFont(weight="bold"))
        self.feed_offset_label.grid(row=2, column=1, sticky="w", pady=4)
        
        self.substrate_dims_label = ctk.CTkLabel(
            grid_frame, 
            text="Substrate Dimensions: -- x -- mm",
            font=ctk.CTkFont(weight="bold")
        )
        self.substrate_dims_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=4)

        # Buttons
        btn_frame = ctk.CTkFrame(sec_calc)
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=12)
        
        ctk.CTkButton(
            btn_frame, 
            text="Calculate Parameters", 
            command=self.calculate_parameters,
            fg_color="#2E8B57", 
            hover_color="#3CB371", 
            width=180
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            btn_frame, 
            text="Save Parameters", 
            command=self.save_parameters,
            fg_color="#4169E1", 
            hover_color="#6495ED", 
            width=140
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            btn_frame, 
            text="Load Parameters", 
            command=self.load_parameters,
            fg_color="#FF8C00", 
            hover_color="#FFA500", 
            width=140
        ).pack(side="left", padx=8)

    def create_tooltip(self, widget, text):
        """Create a tooltip for a widget"""
        def enter(event):
            tooltip = ctk.CTkToplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            label = ctk.CTkLabel(tooltip, text=text, 
                                font=ctk.CTkFont(size=12),
                                fg_color=("gray90", "gray20"),
                                corner_radius=6)
            label.pack()
            widget.tooltip = tooltip
            
        def leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

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
            command=self.start_simulation_thread,
            fg_color="#2E8B57", 
            hover_color="#3CB371", 
            height=40, 
            width=160
        )
        self.run_button.pack(side="left", padx=8)
        
        self.stop_button = ctk.CTkButton(
            btn_row, 
            text="Stop Simulation", 
            command=self.stop_simulation_thread,
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

    def setup_results_tab(self):
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        # Create a paned window for splitting the results area
        paned_window = ctk.CTkFrame(tab)
        paned_window.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        paned_window.grid_columnconfigure(0, weight=1)
        paned_window.grid_rowconfigure(1, weight=1)

        # Title
        ctk.CTkLabel(
            paned_window, 
            text="Results & Beamforming", 
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=10)

        # Graph area with GridSpec 3x2
        graph_frame = ctk.CTkFrame(paned_window)
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
                ax.xaxis.label.set_color('white')
                ax.yaxis.label.set_color('white')
                ax.title.set_color('white')
                for s in ['bottom', 'top', 'right', 'left']:
                    ax.spines[s].set_color('white')
                ax.grid(color='gray', alpha=0.5)
        self.ax_3d.set_facecolor(face)

        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # Beamforming / refresh panel
        panel = ctk.CTkFrame(paned_window)
        panel.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        panel.grid_columnconfigure(0, weight=1)
        self.src_frame = ctk.CTkFrame(panel)
        self.src_frame.grid(row=0, column=0, sticky="ew", padx=6, pady=(8, 2))
        self.source_controls = {}

        ctrl = ctk.CTkFrame(panel)
        ctrl.grid(row=1, column=0, pady=6)
        
        ctk.CTkButton(
            ctrl, 
            text="Analyze S11", 
            command=self.analyze_and_mark_s11,
            fg_color="#6A5ACD", 
            hover_color="#7B68EE"
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            ctrl, 
            text="Apply Sources", 
            command=self.apply_sources_from_ui,
            fg_color="#20B2AA", 
            hover_color="#40E0D0"
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            ctrl, 
            text="Refresh Patterns", 
            command=self.refresh_patterns_only,
            fg_color="#FF8C00", 
            hover_color="#FFA500"
        ).pack(side="left", padx=8)
        
        ctk.CTkCheckBox(
            ctrl, 
            text="Auto-refresh (1.5s)", 
            variable=self.auto_refresh_var,
            command=self.toggle_auto_refresh
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            ctrl, 
            text="Export PNG", 
            command=self.export_png,
            fg_color="#20B2AA", 
            hover_color="#40E0D0"
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            ctrl, 
            text="Export CSV (S11)", 
            command=self.export_csv,
            fg_color="#6A5ACD", 
            hover_color="#7B68EE"
        ).pack(side="left", padx=8)

        self.result_label = ctk.CTkLabel(paned_window, text="", font=ctk.CTkFont(weight="bold"))
        self.result_label.grid(row=3, column=0, pady=(6, 2))

    def setup_log_tab(self):
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            main, 
            text="Simulation Log", 
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=10)

        # Log text area
        self.log_text = ctk.CTkTextbox(
            main, 
            width=900, 
            height=500, 
            font=ctk.CTkFont(family="Consolas")
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")

        # Log buttons
        # Log buttons
        btn_frame = ctk.CTkFrame(main)
        btn_frame.grid(row=2, column=0, pady=8)
        
        ctk.CTkButton(
            btn_frame, 
            text="Clear Log", 
            command=self.clear_log
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            btn_frame, 
            text="Save Log", 
            command=self.save_log
        ).pack(side="left", padx=8)

    def setup_beamforming_tab(self):
        tab = self.tabview.tab("Beamforming")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            main, 
            text="Beamforming Controls", 
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=10)

        # Create a frame for the beamforming controls
        beamforming_frame = ctk.CTkFrame(main)
        beamforming_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        beamforming_frame.grid_columnconfigure(0, weight=1)
        beamforming_frame.grid_rowconfigure(0, weight=1)

        # Add instructions
        instructions = ctk.CTkLabel(
            beamforming_frame,
            text="Adjust amplitude and phase for each port to control the beam pattern.\n"
                 "Changes will be applied without re-running the simulation.",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray70")
        )
        instructions.grid(row=0, column=0, sticky="w", pady=(0, 10))

        # Create a frame for the port controls
        self.port_controls_frame = ctk.CTkFrame(beamforming_frame)
        self.port_controls_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.port_controls_frame.grid_columnconfigure(0, weight=1)

        # Add default message
        self.no_ports_label = ctk.CTkLabel(
            self.port_controls_frame,
            text="No ports available. Run a simulation first.",
            font=ctk.CTkFont(size=14)
        )
        self.no_ports_label.pack(pady=20)

        # Add apply button
        self.apply_beamforming_button = ctk.CTkButton(
            beamforming_frame,
            text="Apply Beamforming Settings",
            command=self.apply_beamforming,
            fg_color="#2E8B57",
            hover_color="#3CB371",
            state="disabled"
        )
        self.apply_beamforming_button.grid(row=2, column=0, pady=10)

    def populate_beamforming_controls(self, port_names):
        """Populate the beamforming controls with sliders for each port"""
        # Clear existing controls
        for widget in self.port_controls_frame.winfo_children():
            widget.destroy()

        if not port_names:
            self.no_ports_label = ctk.CTkLabel(
                self.port_controls_frame,
                text="No ports available.",
                font=ctk.CTkFont(size=14)
            )
            self.no_ports_label.pack(pady=20)
            return

        # Remove the no ports label if it exists
        if hasattr(self, 'no_ports_label'):
            self.no_ports_label.destroy()

        # Create a scrollable frame for the port controls
        scroll_frame = ctk.CTkScrollableFrame(self.port_controls_frame)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        scroll_frame.grid_columnconfigure(0, weight=1)

        # Create headers
        headers = ctk.CTkFrame(scroll_frame)
        headers.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        headers.grid_columnconfigure(1, weight=1)
        headers.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(headers, text="Port", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5)
        ctk.CTkLabel(headers, text="Amplitude", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=5)
        ctk.CTkLabel(headers, text="Phase", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=5)
        ctk.CTkLabel(headers, text="Value", font=ctk.CTkFont(weight="bold")).grid(row=0, column=3, padx=5)

        # Create controls for each port
        self.beamforming_controls = {}
        for i, port_name in enumerate(port_names, 1):
            row_frame = ctk.CTkFrame(scroll_frame)
            row_frame.grid(row=i, column=0, sticky="ew", pady=5)
            row_frame.grid_columnconfigure(1, weight=1)
            row_frame.grid_columnconfigure(3, weight=1)

            # Port name
            ctk.CTkLabel(row_frame, text=port_name).grid(row=0, column=0, padx=5)

            # Amplitude slider
            amp_var = ctk.DoubleVar(value=1.0)
            amp_slider = ctk.CTkSlider(
                row_frame, 
                from_=0, 
                to=2, 
                variable=amp_var,
                number_of_steps=200,
                command=lambda value, p=port_name: self.update_amplitude_value(p, value)
            )
            amp_slider.grid(row=0, column=1, padx=5, sticky="ew")
            
            # Amplitude value display
            amp_value = ctk.CTkLabel(row_frame, text="1.00")
            amp_value.grid(row=0, column=2, padx=5)

            # Phase slider
            phase_var = ctk.DoubleVar(value=0.0)
            phase_slider = ctk.CTkSlider(
                row_frame, 
                from_=0, 
                to=360, 
                variable=phase_var,
                number_of_steps=360,
                command=lambda value, p=port_name: self.update_phase_value(p, value)
            )
            phase_slider.grid(row=0, column=3, padx=5, sticky="ew")
            
            # Phase value display
            phase_value = ctk.CTkLabel(row_frame, text="0.0°")
            phase_value.grid(row=0, column=4, padx=5)

            # Store references to the controls
            self.beamforming_controls[port_name] = {
                'amplitude': amp_slider,
                'amplitude_value': amp_value,
                'amplitude_var': amp_var,
                'phase': phase_slider,
                'phase_value': phase_value,
                'phase_var': phase_var
            }

        # Enable the apply button
        self.apply_beamforming_button.configure(state="normal")

    def update_amplitude_value(self, port_name, value):
        """Update the amplitude value display"""
        value = float(value)
        self.beamforming_controls[port_name]['amplitude_value'].configure(text=f"{value:.2f}")

    def update_phase_value(self, port_name, value):
        """Update the phase value display"""
        value = float(value)
        self.beamforming_controls[port_name]['phase_value'].configure(text=f"{value:.1f}°")

    def apply_beamforming(self):
        """Apply the beamforming settings to the simulation"""
        try:
            if not hasattr(self, 'beamforming_controls') or not self.beamforming_controls:
                self.log_message("No beamforming controls to apply")
                return

            # Get the current excitations
            excitations = self._list_excitations()
            if not excitations:
                self.log_message("No excitations found for beamforming")
                return

            # Prepare the magnitude and phase values
            magnitudes = []
            phases = []
            
            for port_name in excitations:
                if port_name in self.beamforming_controls:
                    amp = self.beamforming_controls[port_name]['amplitude_var'].get()
                    phase = self.beamforming_controls[port_name]['phase_var'].get()
                    magnitudes.append(f"{amp}W")
                    phases.append(f"{phase}deg")
                else:
                    # Use default values if not found
                    magnitudes.append("1W")
                    phases.append("0deg")

            # Apply the changes
            self._edit_sources_with_vars(excitations, magnitudes, phases)
            self.log_message("Beamforming settings applied successfully")
            
            # Refresh the patterns
            self.refresh_patterns_only()
            
        except Exception as e:
            self.log_message(f"Error applying beamforming: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # ------------- Log utilities -------------
    def log_message(self, message):
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
            if self.winfo_exists():
                self.after(100, self.process_log_queue)

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

    def export_csv(self):
        try:
            if hasattr(self, 'simulation_data') and self.simulation_data is not None:
                np.savetxt(
                    "simulation_results.csv", 
                    self.simulation_data, 
                    delimiter=",",
                    header="Frequency (GHz), S11 (dB)", 
                    comments=''
                )
                self.log_message("Data exported to simulation_results.csv")
            else:
                self.log_message("No simulation data available for export")
        except Exception as e:
            self.log_message(f"Error exporting CSV: {e}")

    def export_png(self):
        try:
            if hasattr(self, 'fig'):
                self.fig.savefig("simulation_results.png", dpi=300, bbox_inches='tight')
                self.log_message("Plot saved to simulation_results.png")
        except Exception as e:
            self.log_message(f"Error saving plot: {e}")

    # ----------- Physics / calculations -----------
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
        return True

    def calculate_patch_dimensions(self, frequency_ghz: float) -> Tuple[float, float, float]:
        """Calculate patch dimensions using Balanis formulas for rectangular microstrip patches."""
        f = frequency_ghz * 1e9
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"]) / 1000.0  # mm->m
        
        # Width calculation
        W = self.c / (2 * f) * math.sqrt(2 / (er + 1))
        
        # Effective permittivity
        eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W) ** -0.5
        
        # Length extension due to fringing
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        
        # Effective length
        L_eff = self.c / (2 * f * math.sqrt(eeff))
        
        # Actual length
        L = L_eff - 2 * dL
        
        # Guided wavelength
        lambda_g = self.c / (f * math.sqrt(eeff))
        
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)  # Convert to mm

    def _size_array_from_gain(self):
        """Calculate array size based on desired gain."""
        G_elem = 8.0  # dBi (typical gain for a single patch)
        G_des = float(self.params["gain"])
        
        # Calculate required number of elements
        N_req = max(1, int(math.ceil(10 ** ((G_des - G_elem) / 10.0))))
        
        # Ensure even number of elements
        if N_req % 2 == 1:
            N_req += 1
            
        # Calculate rows and columns
        rows = max(2, int(round(math.sqrt(N_req))))
        if rows % 2 == 1:
            rows += 1
            
        cols = max(2, int(math.ceil(N_req / rows)))
        if cols % 2 == 1:
            cols += 1
            
        # Ensure we have enough elements
        while rows * cols < N_req:
            if rows <= cols:
                rows += 2
            else:
                cols += 2
                
        return rows, cols, N_req

    def calculate_substrate_size(self):
        """Calculate substrate dimensions based on array configuration."""
        L = self.calculated_params["patch_length"]
        W = self.calculated_params["patch_width"]
        s = self.calculated_params["spacing"]
        r = self.calculated_params["rows"]
        c = self.calculated_params["cols"]
        
        total_w = c * W + (c - 1) * s
        total_l = r * L + (r - 1) * s
        
        # Add margin (20% of the larger dimension)
        margin = max(total_w, total_l) * 0.20
        
        self.calculated_params["substrate_width"] = total_w + 2 * margin
        self.calculated_params["substrate_length"] = total_l + 2 * margin
        
        self.log_message(
            f"Substrate size calculated: {self.calculated_params['substrate_width']:.2f} x "
            f"{self.calculated_params['substrate_length']:.2f} mm"
        )

    def calculate_parameters(self):
        """Calculate all antenna parameters based on user inputs."""
        self.log_message("Starting parameter calculation")
        
        if not self.get_parameters():
            self.log_message("Parameter calculation failed due to invalid input")
            return
            
        try:
            # Calculate patch dimensions
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params.update({
                "patch_length": L_mm, 
                "patch_width": W_mm, 
                "lambda_g": lambda_g_mm
            })
            
            # Calculate spacing
            lambda0_m = self.c / (self.params["frequency"] * 1e9)
            factors = {
                "lambda/2": 0.5, 
                "lambda": 1.0, 
                "0.7*lambda": 0.7, 
                "0.8*lambda": 0.8, 
                "0.9*lambda": 0.9
            }
            
            spacing_mm = factors.get(self.params["spacing_type"], 0.5) * lambda0_m * 1000.0
            self.calculated_params["spacing"] = spacing_mm
            
            # Calculate number of elements
            rows, cols, N_req = self._size_array_from_gain()
            self.calculated_params.update({
                "num_patches": rows * cols, 
                "rows": rows, 
                "cols": cols
            })
            
            self.log_message(
                f"Array sizing -> target gain {self.params['gain']} dBi, N_req≈{N_req}, "
                f"layout {rows}x{cols} (= {rows*cols} patches)"
            )
            
            # Calculate feed offset in y direction
            self.calculated_params["feed_offset"] = 0.30 * L_mm
            
            # Calculate substrate dimensions
            self.calculate_substrate_size()
            
            # Update UI with calculated values
            self.patches_label.configure(text=f"Number of Patches: {rows*cols}")
            self.rows_cols_label.configure(text=f"Configuration: {rows} x {cols}")
            self.spacing_label.configure(text=f"Spacing: {spacing_mm:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch Dimensions: {L_mm:.2f} x {W_mm:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {lambda_g_mm:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset (y): {self.calculated_params['feed_offset']:.2f} mm")
            
            self.substrate_dims_label.configure(
                text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x "
                     f"{self.calculated_params['substrate_length']:.2f} mm"
            )
            
            self.status_label.configure(text="Parameters calculated successfully")
            self.log_message("Parameters calculated successfully")
            
        except Exception as e:
            self.status_label.configure(text=f"Error in calculation: {e}")
            self.log_message(f"Error in calculation: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # --------- AEDT helpers ---------
    def _ensure_material(self, name: str, er: float, tan_d: float):
        """Ensure a material exists in the project, create it if not."""
        try:
            if not self.hfss.materials.checkifmaterialexists(name):
                self.hfss.materials.add_material(name)
                m = self.hfss.materials.material_keys[name]
                m.permittivity = er
                m.dielectric_loss_tangent = tan_d
                self.log_message(f"Created material: {name} (er={er}, tanδ={tan_d})")
        except Exception as e:
            self.log_message(f"Material management warning for '{name}': {e}")

    def _open_or_create_project(self):
        """Open an existing project or create a new one."""
        if self.desktop is None:
            self.desktop = Desktop(
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"],
                new_desktop=True
            )

        # Try to find existing project
        od = getattr(self.desktop, "_odesktop", None) or getattr(self.desktop, "odesktop", None)
        open_names, open_objs = [], []
        
        if od:
            try:
                for p in od.GetProjects():
                    try:
                        open_names.append(p.GetName())
                        open_objs.append(p)
                    except Exception:
                        pass
            except Exception:
                pass

        # Use existing project if found
        if self.project_display_name in open_names:
            idx = open_names.index(self.project_display_name)
            proj_obj = open_objs[idx]
            new_design = self.design_base_name
            
            try:
                tmp = Hfss(
                    project=proj_obj, 
                    non_graphical=self.params["non_graphical"],
                    version=self.params["aedt_version"]
                )
                
                existing = list(tmp.project.design_list)
                k = 1
                while new_design in existing:
                    k += 1
                    new_design = f"{self.design_base_name}_{k}"
                    
                try:
                    tmp.close_desktop()
                except Exception:
                    pass
                    
            except Exception:
                new_design = f"{self.design_base_name}_{datetime.now().strftime('%H%M%S')}"
                
            self.hfss = Hfss(
                project=proj_obj, 
                design=new_design, 
                solution_type="DrivenModal",
                version=self.params["aedt_version"], 
                non_graphical=self.params["non_graphical"]
            )
            
            try:
                self.project_path = proj_obj.GetPath()
            except Exception:
                self.project_path = ""
                
            self.log_message(f"Using existing project '{self.project_display_name}', created design '{new_design}'")
            return

        # Create new project
        if self.temp_folder is None:
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
            
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.project_path = os.path.join(self.temp_folder.name, f"{self.project_display_name}_{ts}.aedt")
        
        self.hfss = Hfss(
            project=self.project_path, 
            design=self.design_base_name, 
            solution_type="DrivenModal",
            version=self.params["aedt_version"], 
            non_graphical=self.params["non_graphical"]
        )
        
        self.log_message(f"Created new project: {self.project_path} (design '{self.design_base_name}')")

    def _set_design_variables(self, L, W, spacing, rows, cols, h_sub, sub_w, sub_l):
        """Set design variables in HFSS."""
        a = float(self.params["probe_radius"])
        ba = float(self.params["coax_ba_ratio"])
        b = a * ba
        wall = float(self.params["coax_wall_thickness"])
        Lp = float(self.params["coax_port_length"])
        clear = float(self.params["antipad_clearance"])

        # Set variables
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
        self.hfss["padAir"] = f"{max(spacing, W, L)/2 + Lp + 2.0}mm"
        
        self.log_message(f"Air coax set: a={a:.3f} mm, b={b:.3f} mm (b/a={ba:.3f}≈2.3 → ~50 Ω)")
        
        return a, b, wall, Lp, clear

    def _create_coax_feed_lumped(self, ground, substrate, x_feed: float, y_feed: float, name_prefix: str):
        """Create a coaxial feed with lumped port."""
        try:
            a_val = float(self.params["probe_radius"])
            b_val = a_val * float(self.params["coax_ba_ratio"])
            wall_val = float(self.params["coax_wall_thickness"])
            Lp_val = float(self.params["coax_port_length"])
            h_sub_val = float(self.params["substrate_thickness"])
            clear_val = float(self.params["antipad_clearance"])
            
            # Ensure minimum gap between inner and outer conductor
            if b_val - a_val < 0.02:
                b_val = a_val + 0.02

            # Create pin
            pin = self.hfss.modeler.create_cylinder(
                orientation="Z", 
                origin=[x_feed, y_feed, -Lp_val],
                radius=a_val, 
                height=h_sub_val + Lp_val + 0.001,
                name=f"{name_prefix}_Pin", 
                material="copper"
            )

            # Create shield (outer conductor)
            shield_outer = self.hfss.modeler.create_cylinder(
                orientation="Z", 
                origin=[x_feed, y_feed, -Lp_val],
                radius=b_val + wall_val, 
                height=Lp_val,
                name=f"{name_prefix}_ShieldOuter", 
                material="copper"
            )
            
            # Create void for shield
            shield_inner_void = self.hfss.modeler.create_cylinder(
                orientation="Z", 
                origin=[x_feed, y_feed, -Lp_val],
                radius=b_val, 
                height=Lp_val,
                name=f"{name_prefix}_ShieldInnerVoid", 
                material="vacuum"
            )
            
            # Subtract void from shield
            self.hfss.modeler.subtract(shield_outer, [shield_inner_void], keep_originals=False)

            # Create hole in substrate
            hole_r = b_val + clear_val
            sub_hole = self.hfss.modeler.create_cylinder(
                orientation="Z", 
                origin=[x_feed, y_feed, 0.0],
                radius=hole_r, 
                height=h_sub_val, 
                name=f"{name_prefix}_SubHole", 
                material="vacuum"
            )
            
            # Subtract hole from substrate
            self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=False)
            
            # Create hole in ground plane
            g_hole = self.hfss.modeler.create_circle(
                orientation="XY", 
                origin=[x_feed, y_feed, 0.0], 
                radius=hole_r,
                name=f"{name_prefix}_GndHole", 
                material="vacuum"
            )
            
            # Subtract hole from ground
            self.hfss.modeler.subtract(ground, [g_hole], keep_originals=False)

            # Create port ring
            port_ring = self.hfss.modeler.create_circle(
                orientation="XY", 
                origin=[x_feed, y_feed, -Lp_val],
                radius=b_val, 
                name=f"{name_prefix}_PortRing", 
                material="vacuum"
            )
            
            # Create port hole
            port_hole = self.hfss.modeler.create_circle(
                orientation="XY", 
                origin=[x_feed, y_feed, -Lp_val],
                radius=a_val, 
                name=f"{name_prefix}_PortHole", 
                material="vacuum"
            )
            
            # Subtract hole from ring
            self.hfss.modeler.subtract(port_ring, [port_hole], keep_originals=False)

            # Create integration line for port
            eps_line = min(0.1 * (b_val - a_val), 0.05)
            r_start = a_val + eps_line
            r_end = b_val - eps_line
            
            if r_end <= r_start:
                r_end = a_val + 0.75 * (b_val - a_val)
                
            p1 = [x_feed + r_start, y_feed, -Lp_val]
            p2 = [x_feed + r_end,   y_feed, -Lp_val]

            # Create lumped port
            port_name = f"{name_prefix}_Lumped"
            _ = self.hfss.lumped_port(
                assignment=port_ring.name,
                integration_line=[p1, p2],
                impedance=50.0,
                name=port_name,
                renormalize=True
            )
            
            # Add to created ports list
            if port_name not in self.created_ports:
                self.created_ports.append(port_name)
            
            self.log_message(f"Lumped Port '{port_name}' created (integration line).")
            return pin, None, shield_outer
            
        except Exception as e:
            self.log_message(f"Exception in coax creation '{name_prefix}': {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            return None, None, None

    # ---------- Post-solve helpers ----------
# Modificações na classe ModernPatchAntennaDesigner

    def _add_post_var(self, name: str, value: str) -> bool:
        """Add a post-processing variable to the design"""
        try:
            # First check if variable already exists
            try:
                existing_value = self.hfss.odesign.GetVariableValue(name)
                self.log_message(f"Post var '{name}' already exists with value: {existing_value}")
                return True
            except:
                pass
            
            # Try to create the variable using the COM interface
            self.hfss.odesign.CreateVariable(name, value)
            self.log_message(f"Post var '{name}' = {value} created successfully.")
            return True
        except Exception as e:
            self.log_message(f"Add post var '{name}' failed: {e}")
            # Try alternative method using variable manager
            try:
                self.hfss.variable_manager.set_variable(name, value)
                self.log_message(f"Post var '{name}' = {value} created via variable_manager.")
                return True
            except Exception as e2:
                self.log_message(f"Add post var '{name}' also failed via variable_manager: {e2}")
                return False

    def _edit_sources_with_vars(self, excitations: List[str], magnitudes: List[str], phases: List[str]) -> bool:
        """Edit sources with post-processing variables"""
        try:
            # Get the Solutions module
            sol = self.hfss.odesign.GetModule("Solutions")
            
            # Prepare the argument list for EditSources
            args = ["IncludePortPostProcessing:=", False, "SpecifySystemPower:=", False]
            
            sources = []
            for ex, mag, ph in zip(excitations, magnitudes, phases):
                sources.append(["Name:=", ex, "Magnitude:=", mag, "Phase:=", ph])
            
            # Combine the arguments and sources
            full_args = args + sources
            
            # Execute EditSources
            sol.EditSources(full_args)
            self.log_message(f"EditSources applied to {len(excitations)} port(s).")
            return True
        except Exception as e:
            self.log_message(f"EditSources failed: {e}")
            # Try alternative approach using the excitation object
            try:
                for i, (ex, mag, ph) in enumerate(zip(excitations, magnitudes, phases)):
                    # Get the excitation object
                    excitation = self.hfss.excitations[ex]
                    # Set magnitude and phase
                    excitation.magnitude = mag
                    excitation.phase = ph
                self.log_message(f"Excitations updated directly for {len(excitations)} port(s).")
                return True
            except Exception as e2:
                self.log_message(f"Direct excitation update also failed: {e2}")
                return False

    def ensure_post_processing_vars(self):
        """Ensure post-processing variables exist before trying to use them"""
        try:
            excitations = self._list_excitations()
            if not excitations:
                self.log_message("No excitations found for post-processing variables.")
                return False
                
            for i in range(1, len(excitations) + 1):
                p_name = f"p{i}"
                ph_name = f"ph{i}"
                
                # Check if variables exist, create if not
                try:
                    self.hfss.odesign.GetVariableValue(p_name)
                    self.log_message(f"Variable {p_name} already exists.")
                except:
                    self._add_post_var(p_name, "1W")
                    
                try:
                    self.hfss.odesign.GetVariableValue(ph_name)
                    self.log_message(f"Variable {ph_name} already exists.")
                except:
                    self._add_post_var(ph_name, "0deg")
                    
            return True
        except Exception as e:
            self.log_message(f"Error ensuring post-processing variables: {e}")
            return False

    def _postprocess_after_solve(self):
        """Post-processing after simulation is complete"""
        try:
            # Get excitations
            excitations = self._list_excitations()
            if not excitations:
                self.log_message("No excitations found for post-processing.")
                return

            # Ensure post-processing variables exist
            self.ensure_post_processing_vars()
            
            # Create post-processing variables
            pvars, phvars = [], []
            for i in range(1, len(excitations) + 1):
                p_name = f"p{i}"
                ph_name = f"ph{i}"
                
                # Create variables if they don't exist
                self._add_post_var(p_name, "1W")
                self._add_post_var(ph_name, "0deg")
                
                pvars.append(p_name)
                phvars.append(ph_name)

            # Apply the variables to the sources
            success = self._edit_sources_with_vars(excitations, pvars, phvars)
            
            if success:
                # Populate beamforming controls
                self.populate_beamforming_controls(excitations)
                
                # Force a refresh of the solution data
                try:
                    self.hfss.post.clear_solutions()
                    self.hfss.post.load_solution_data()
                except Exception as e:
                    self.log_message(f"Warning: Could not refresh solution data: {e}")
                
                self.log_message("Post-processing completed successfully")
            else:
                self.log_message("Post-processing completed with warnings - sources not updated")
                
        except Exception as e:
            self.log_message(f"Error in post-processing: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    def refresh_patterns_only(self):
        """Refresh radiation patterns without re-running simulation"""
        try:
            # Ensure post-processing variables exist
            if not self.ensure_post_processing_vars():
                self.log_message("Cannot refresh patterns - post-processing variables not available")
                return
                
            self.log_message("Refreshing radiation patterns")
            
            # Clear radiation pattern axes
            self.ax_th.clear()
            self.ax_ph.clear()
            self.ax_3d.clear()
            
            f0 = float(self.params["frequency"])

            # Theta cut (Phi = 0°)
            th, gth = self._get_gain_cut(f0, cut="theta", fixed_angle_deg=0.0)
            if th is not None and gth is not None:
                self.theta_cut = (th, gth)
                self.ax_th.plot(th, gth, linewidth=2)
                self.ax_th.set_xlabel("Theta (deg)")
                self.ax_th.set_ylabel("Gain (dB)")
                self.ax_th.set_title("Radiation Pattern - Theta cut (Phi = 0°)")
                self.ax_th.grid(True, alpha=0.5)
            else:
                self.ax_th.text(0.5, 0.5, "Theta-cut gain not available",
                                transform=self.ax_th.transAxes, ha="center", va="center")

            # Phi cut (Theta = 90°)
            ph, gph = self._get_gain_cut(f0, cut="phi", fixed_angle_deg=90.0)
            if ph is not None and gph is not None:
                self.phi_cut = (ph, gph)
                self.ax_ph.plot(ph, gph, linewidth=2)
                self.ax_ph.set_xlabel("Phi (deg)")
                self.ax_ph.set_ylabel("Gain (dB)")
                self.ax_ph.set_title("Radiation Pattern - Phi cut (Theta = 90°)")
                self.ax_ph.grid(True, alpha=0.5)
            else:
                self.ax_ph.text(0.5, 0.5, "Phi-cut gain not available",
                                transform=self.ax_ph.transAxes, ha="center", va="center")

            # Redraw the canvas
            self.canvas.draw()
            self.log_message("Radiation patterns refreshed")
            
        except Exception as e:
            self.log_message(f"Error refreshing patterns: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    def _ensure_infinite_sphere(self, name="Infinite Sphere1"):
        """Create an infinite sphere for far-field analysis."""
        try:
            rf = self.hfss.odesign.GetModule("RadField")
            props = [
                f"NAME:{name}",
                "UseCustomRadiationSurface:=", False,
                "CSDefinition:=", "Theta-Phi",
                "Polarization:=", "Linear",
                "ThetaStart:=", "-180deg",
                "ThetaStop:=", "180deg",
                "ThetaStep:=", "1deg",
                "PhiStart:=", "-180deg",
                "PhiStop:=", "180deg",
                "PhiStep:=", "1deg",
                "UseLocalCS:=", False
            ]
            
            rf.InsertInfiniteSphereSetup(props)
            self.log_message(f"Infinite sphere '{name}' created (theta -180..180, phi -180..180, 1deg).")
            return name
            
        except Exception as e:
            self.log_message(f"Infinite sphere creation failed: {e}")
            return None

    def ensure_post_processing_vars(self):
        """Ensure post-processing variables exist before trying to use them"""
        try:
            excitations = self._list_excitations()
            if not excitations:
                return False
                
            for i in range(1, len(excitations) + 1):
                p_name = f"p{i}"
                ph_name = f"ph{i}"
                
                # Check if variables exist, create if not
                try:
                    self.hfss.odesign.GetVariableValue(p_name)
                except:
                    self._add_post_var(p_name, "1W")
                    
                try:
                    self.hfss.odesign.GetVariableValue(ph_name)
                except:
                    self._add_post_var(ph_name, "0deg")
                    
            return True
        except Exception as e:
            self.log_message(f"Error ensuring post-processing variables: {e}")
            return False

    def _postprocess_after_solve(self):
        """Post-processing after simulation is complete"""
        try:
            # Get excitations
            excitations = self._list_excitations()
            if not excitations:
                self.log_message("No excitations found for post-processing.")
                return

            # Create post-processing variables
            pvars, phvars = [], []
            for i in range(1, len(excitations) + 1):
                p_name = f"p{i}"
                ph_name = f"ph{i}"
                
                # Create variables if they don't exist
                self._add_post_var(p_name, "1W")
                self._add_post_var(ph_name, "0deg")
                
                pvars.append(p_name)
                phvars.append(ph_name)

            # Apply the variables to the sources
            success = self._edit_sources_with_vars(excitations, pvars, phvars)
            
            if success:
                # Populate beamforming controls
                self.populate_beamforming_controls(excitations)
                
                # Force a refresh of the solution data
                try:
                    self.hfss.post.clear_solutions()
                    self.hfss.post.load_solution_data()
                except Exception as e:
                    self.log_message(f"Warning: Could not refresh solution data: {e}")
                
                self.log_message("Post-processing completed successfully")
            else:
                self.log_message("Post-processing completed with warnings - sources not updated")
                
        except Exception as e:
            self.log_message(f"Error in post-processing: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # ------------- Far Field Analysis -------------
    def _get_gain_cut(self, frequency: float, cut: str, fixed_angle_deg: float):
        """Get far-field gain data for a specific cut (theta or phi) at a given frequency."""
        try:
            # Create a unique report name based on the cut
            report_name = f"Gain_{cut}_cut_{fixed_angle_deg}deg"
            
            # Define the expression for gain in dB
            expression = "dB(GainTotal)"
            
            # Set up variations based on cut type
            if cut == "theta":
                # Theta cut: vary theta, fix phi
                variations = {
                    "Freq": [frequency, "GHz"],
                    "Phi": [fixed_angle_deg, "deg"],
                    "Theta": ["All"]
                }
                primary_sweep = "Theta"
            else:
                # Phi cut: vary phi, fix theta
                variations = {
                    "Freq": [frequency, "GHz"],
                    "Theta": [fixed_angle_deg, "deg"],
                    "Phi": ["All"]
                }
                primary_sweep = "Phi"
            
            # Try to get existing report or create new one
            try:
                report = self.hfss.post.reports_by_category.far_field(
                    expressions=[expression],
                    context="Infinite Sphere1",
                    primary_sweep_variable=primary_sweep,
                    variations=variations,
                    name=report_name
                )
            except:
                # If report doesn't exist, create it
                report = self.hfss.post.create_report(
                    expressions=[expression],
                    context="Infinite Sphere1",
                    primary_sweep=primary_sweep,
                    variations=variations,
                    report_category="Far Fields",
                    name=report_name
                )
            
            # Get the solution data
            solution_data = report.get_solution_data()
            
            if solution_data and hasattr(solution_data, 'primary_sweep_values'):
                angles = np.array(solution_data.primary_sweep_values)
                gains = np.array(solution_data.data_real())[0]  # Get first (and only) expression data
                
                # Ensure we have valid data
                if len(angles) > 0 and len(angles) == len(gains):
                    return angles, gains
            
            return None, None
            
        except Exception as e:
            self.log_message(f"Error getting {cut} cut data: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            return None, None

    def _get_gain_3d_grid(self, frequency: float, theta_step=10.0, phi_step=10.0):
        """Get 3D gain data for visualization"""
        try:
            # Create theta and phi arrays
            theta_vals = np.arange(-180, 180 + theta_step, theta_step)
            phi_vals = np.arange(-180, 180 + phi_step, phi_step)
            
            # Initialize gain matrix
            gain_matrix = np.zeros((len(theta_vals), len(phi_vals)))
            
            # Get gain values for each theta-phi combination
            for i, theta in enumerate(theta_vals):
                for j, phi in enumerate(phi_vals):
                    # Get gain for this specific direction
                    expression = "dB(GainTotal)"
                    variations = {
                        "Freq": [frequency, "GHz"],
                        "Theta": [theta, "deg"],
                        "Phi": [phi, "deg"]
                    }
                    
                    report = self.hfss.post.reports_by_category.far_field(
                        expressions=[expression],
                        context="Infinite Sphere1",
                        variations=variations
                    )
                    
                    solution_data = report.get_solution_data()
                    if solution_data and hasattr(solution_data, 'data_real'):
                        gain = float(solution_data.data_real()[0])
                        gain_matrix[i, j] = gain
            
            return theta_vals, phi_vals, gain_matrix
            
        except Exception as e:
            self.log_message(f"Error getting 3D gain data: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            return None, None, None

    # ------------- Simulation -------------
    def start_simulation_thread(self):
        """Start simulation in a separate thread."""
        if self.is_simulation_running:
            self.log_message("Simulation is already running")
            return
            
        self.stop_simulation = False
        self.is_simulation_running = True
        
        threading.Thread(target=self.run_simulation, daemon=True).start()

    def stop_simulation_thread(self):
        """Request simulation stop."""
        self.stop_simulation = True
        self.log_message("Simulation stop requested")

    def run_simulation(self):
        """Run the simulation."""
        try:
            self.log_message("Starting simulation")
            self.run_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            self.sim_status_label.configure(text="Simulation in progress")
            self.progress_bar.set(0)

            # Get parameters
            if not self.get_parameters():
                self.log_message("Invalid parameters. Aborting.")
                return
                
            # Calculate parameters if not already done
            if self.calculated_params["num_patches"] < 1:
                self.calculate_parameters()

            # Open or create project
            self._open_or_create_project()
            self.progress_bar.set(0.25)

            # Set model units
            self.hfss.modeler.model_units = "mm"
            self.log_message("Model units set to: mm")

            # Handle substrate material
            sub_name = self.params["substrate_material"]
            if not self.hfss.materials.checkifmaterialexists(sub_name):
                sub_name = "Custom_Substrate"
                self._ensure_material(sub_name, float(self.params["er"], float(self.params["tan_d"])))

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
            self.log_message(f"Creating {rows*cols} patches in {rows*cols} configuration")
            patches = []
            total_w = cols * W + (cols - 1) * spacing
            total_l = rows * L + (rows - 1) * spacing
            start_x = -total_w / 2 + W / 2
            start_y = -total_l / 2 + L / 2
            self.progress_bar.set(0.35)

            # Create each patch
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
                    self.progress_bar.set(0.35 + 0.25 * (count / float(rows * cols)))

            # Check if simulation was stopped
            if self.stop_simulation:
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
            self.progress_bar.set(0.65)

            # Create infinite sphere for far-field analysis
            self._ensure_infinite_sphere("Infinite Sphere1")

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
                
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])

            # Check if simulation was stopped
            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return

            # Perform post-processing
            self._postprocess_after_solve()

            # Process results
            self.progress_bar.set(0.9)
            self.log_message("Processing results")
            self.plot_results()
            
            # Finalize
            self.progress_bar.set(1.0)
            self.sim_status_label.configure(text="Simulation completed")
            self.log_message("Simulation completed successfully")
            
        except Exception as e:
            self.log_message(f"Error in simulation: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            self.sim_status_label.configure(text=f"Simulation error: {e}")
            
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    def _list_excitations(self):
        """Get list of excitations in the design"""
        try:
            # Try to get excitations from HFSS
            excitations = self.hfss.get_excitations_name()
            if excitations:
                return excitations
        except Exception:
            pass
        
        # Fallback to created ports list
        return [f"{port}:1" for port in self.created_ports]

    def plot_results(self):
        """Plot simulation results."""
        try:
            self.log_message("Plotting results")

            # Clear all axes
            self.ax_s11.clear()
            self.ax_imp.clear()
            self.ax_th.clear()
            self.ax_ph.clear()
            self.ax_3d.clear()
            
            # Determine S-parameter expression
            excitations = self._list_excitations()
            expr = "dB(S(1,1))"
            if excitations:
                # Try to get the first port name correctly
                try:
                    p = excitations[0].split(":")[0]
                    expr = f"dB(S({p},{p}))"
                except:
                    # Fallback to general S11
                    expr = "dB(S(1,1))"

            # Get S-parameter data
            try:
                # Try different methods to get S-parameter data
                try:
                    rpt = self.hfss.post.reports_by_category.standard(expressions=[expr])
                    rpt.context = ["Setup1: Sweep1"]
                    sol = rpt.get_solution_data()
                except:
                    # Alternative method
                    sol = self.hfss.post.get_solution_data(
                        expressions=[expr],
                        context="Sweep1",
                        report_category="Standard"
                    )
            except Exception as e:
                self.log_message(f"Error getting S-parameter data: {e}")
                sol = None

            # Plot S11
            if sol and hasattr(sol, 'primary_sweep_values') and hasattr(sol, 'data_real'):
                try:
                    freqs = np.asarray(sol.primary_sweep_values, dtype=float)
                    data = sol.data_real()
                    
                    if isinstance(data, (list, tuple)) and len(data) > 0 and hasattr(data[0], "__len__"):
                        y = np.asarray(data[0], dtype=float)
                    else:
                        y = np.asarray(data, dtype=float)
                        
                    if y.size == freqs.size:
                        self.simulation_data = np.column_stack((freqs, y))
                        
                        # Plot S11
                        self.ax_s11.plot(freqs, y, label='S11', linewidth=2)
                        self.ax_s11.axhline(y=-10, linestyle='--', alpha=0.7, label='-10 dB')
                        self.ax_s11.set_xlabel("Frequency (GHz)")
                        self.ax_s11.set_ylabel("S-Parameter (dB)")
                        self.ax_s11.set_title("S11 Parameter")
                        self.ax_s11.legend()
                        self.ax_s11.grid(True, alpha=0.5)
                        
                        # Mark center frequency
                        cf = float(self.params["frequency"])
                        self.ax_s11.axvline(x=cf, linestyle='--', alpha=0.7, color='red')
                        self.ax_s11.text(cf, self.ax_s11.get_ylim()[1]*0.9, f'{cf} GHz', color='red')
                        
                except Exception as e:
                    self.log_message(f"S11 plotting warning: {e}")
            else:
                self.ax_s11.text(0.5, 0.5, "S11 data not available",
                                transform=self.ax_s11.transAxes, ha="center", va="center")

            # ---------- Far-Field cuts ----------
            f0 = float(self.params["frequency"])

            # Theta cut (Phi = 0°)
            th, gth = self._get_gain_cut(f0, cut="theta", fixed_angle_deg=0.0)
            if th is not None and gth is not None and len(th) == len(gth) and len(th) > 0:
                self.theta_cut = (th, gth)
                self.ax_th.plot(th, gth, linewidth=2)
                self.ax_th.set_xlabel("Theta (deg)")
                self.ax_th.set_ylabel("Gain (dB)")
                self.ax_th.set_title("Radiation Pattern - Theta cut (Phi = 0°)")
                self.ax_th.grid(True, alpha=0.5)
            else:
                self.ax_th.text(0.5, 0.5, "Theta-cut gain not available",
                                transform=self.ax_th.transAxes, ha="center", va="center")

            # Phi cut (Theta = 90°)
            ph, gph = self._get_gain_cut(f0, cut="phi", fixed_angle_deg=90.0)
            if ph is not None and gph is not None and len(ph) == len(gph) and len(ph) > 0:
                self.phi_cut = (ph, gph)
                self.ax_ph.plot(ph, gph, linewidth=2)
                self.ax_ph.set_xlabel("Phi (deg)")
                self.ax_ph.set_ylabel("Gain (dB)")
                self.ax_ph.set_title("Radiation Pattern - Phi cut (Theta = 90°)")
                self.ax_ph.grid(True, alpha=0.5)
            else:
                self.ax_ph.text(0.5, 0.5, "Phi-cut gain not available",
                                transform=self.ax_ph.transAxes, ha="center", va="center")

            # 3D radiation pattern
            try:
                # Get 3D gain data
                theta, phi, gain = self._get_gain_3d_grid(f0, 
                                                         theta_step=self.params["theta_step"],
                                                         phi_step=self.params["phi_step"])
                
                if theta is not None and phi is not None and gain is not None:
                    # Convert to radians for 3D plotting
                    theta_rad = np.radians(theta)
                    phi_rad = np.radians(phi)
                    
                    # Create meshgrid
                    TH, PH = np.meshgrid(theta_rad, phi_rad)
                    
                    # Convert gain to linear scale for visualization
                    gain_linear = 10**(gain/10)
                    
                    # Normalize for better visualization
                    gain_normalized = gain_linear / np.max(gain_linear)
                    
                    # Convert to Cartesian coordinates
                    X = gain_normalized * np.sin(TH) * np.cos(PH)
                    Y = gain_normalized * np.sin(TH) * np.sin(PH)
                    Z = gain_normalized * np.cos(TH)
                    
                    # Plot 3D surface
                    surf = self.ax_3d.plot_surface(X, Y, Z, cmap=cm.viridis, alpha=0.8, linewidth=0)
                    self.ax_3d.set_title("3D Radiation Pattern")
                    self.fig.colorbar(surf, ax=self.ax_3d, shrink=0.5, aspect=20, label='Gain (dB)')
                    
            except Exception as e:
                self.log_message(f"3D pattern plotting warning: {e}")
                self.ax_3d.text(0.5, 0.5, 0.5, "3D pattern not available", 
                               transform=self.ax_3d.transAxes, ha="center", va="center")

            # Finalize plot
            self.fig.tight_layout()
            self.canvas.draw()
            self.log_message("Results plotted successfully")
            
        except Exception as e:
            self.log_message(f"Error plotting results: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    def analyze_and_mark_s11(self):
        """Analyze S11 and mark key points - enhanced version"""
        try:
            if self.simulation_data is None:
                # Try to get S11 data directly from HFSS if not available
                try:
                    expr = "dB(S(1,1))"
                    excitations = self._list_excitations()
                    if excitations:
                        try:
                            p = excitations[0].split(":")[0]
                            expr = f"dB(S({p},{p}))"
                        except:
                            pass
                    
                    sol = self.hfss.post.get_solution_data(
                        expressions=[expr],
                        context="Sweep1",
                        report_category="Standard"
                    )
                    
                    if sol and hasattr(sol, 'primary_sweep_values') and hasattr(sol, 'data_real'):
                        freqs = np.asarray(sol.primary_sweep_values, dtype=float)
                        data = sol.data_real()
                        
                        if isinstance(data, (list, tuple)) and len(data) > 0 and hasattr(data[0], "__len__"):
                            y = np.asarray(data[0], dtype=float)
                        else:
                            y = np.asarray(data, dtype=float)
                            
                        if y.size == freqs.size:
                            self.simulation_data = np.column_stack((freqs, y))
                except Exception as e:
                    self.log_message(f"Error retrieving S11 data for analysis: {e}")
                    return
                    
            if self.simulation_data is None:
                self.log_message("No simulation data available for analysis")
                return
                
            # Find the frequency with minimum S11 (peak resonance)
            frequencies = self.simulation_data[:, 0]
            s11 = self.simulation_data[:, 1]
            min_idx = np.argmin(s11)
            resonant_freq = frequencies[min_idx]
            min_s11 = s11[min_idx]
            
            # Mark the resonant frequency on the plot
            self.ax_s11.plot(resonant_freq, min_s11, 'ro', markersize=8)
            self.ax_s11.annotate(f'Resonance: {resonant_freq:.3f} GHz\nS11: {min_s11:.2f} dB',
                                xy=(resonant_freq, min_s11),
                                xytext=(resonant_freq + 0.5, min_s11 + 5),
                                arrowprops=dict(facecolor='black', shrink=0.05))
            
            # Calculate bandwidth at -10 dB
            below_10db = np.where(s11 <= -10)[0]
            if len(below_10db) > 0:
                bw_start = frequencies[below_10db[0]]
                bw_end = frequencies[below_10db[-1]]
                bandwidth = bw_end - bw_start
                
                # Mark bandwidth on the plot
                self.ax_s11.axvline(x=bw_start, color='green', linestyle='--', alpha=0.7)
                self.ax_s11.axvline(x=bw_end, color='green', linestyle='--', alpha=0.7)
                self.ax_s11.text((bw_start + bw_end)/2, self.ax_s11.get_ylim()[0] + 5,
                                f'Bandwidth: {bandwidth:.3f} GHz', 
                                ha='center', color='green')
            
            self.canvas.draw()
            self.log_message(f"S11 analysis complete: Resonance at {resonant_freq:.3f} GHz")
            
        except Exception as e:
            self.log_message(f"Error in S11 analysis: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    def refresh_patterns_only(self):
        """Refresh radiation patterns without re-running simulation"""
        try:
            # Ensure post-processing variables exist
            self.ensure_post_processing_vars()
            
            self.log_message("Refreshing radiation patterns")
            
            # Clear radiation pattern axes
            self.ax_th.clear()
            self.ax_ph.clear()
            self.ax_3d.clear()
            
            f0 = float(self.params["frequency"])

            # Theta cut (Phi = 0°)
            th, gth = self._get_gain_cut(f0, cut="theta", fixed_angle_deg=0.0)
            if th is not None and gth is not None:
                self.theta_cut = (th, gth)
                self.ax_th.plot(th, gth, linewidth=2)
                self.ax_th.set_xlabel("Theta (deg)")
                self.ax_th.set_ylabel("Gain (dB)")
                self.ax_th.set_title("Radiation Pattern - Theta cut (Phi = 0°)")
                self.ax_th.grid(True, alpha=0.5)
            else:
                self.ax_th.text(0.5, 0.5, "Theta-cut gain not available",
                                transform=self.ax_th.transAxes, ha="center", va="center")

            # Phi cut (Theta = 90°)
            ph, gph = self._get_gain_cut(f0, cut="phi", fixed_angle_deg=90.0)
            if ph is not None and gph is not None:
                self.phi_cut = (ph, gph)
                self.ax_ph.plot(ph, gph, linewidth=2)
                self.ax_ph.set_xlabel("Phi (deg)")
                self.ax_ph.set_ylabel("Gain (dB)")
                self.ax_ph.set_title("Radiation Pattern - Phi cut (Theta = 90°)")
                self.ax_ph.grid(True, alpha=0.5)
            else:
                self.ax_ph.text(0.5, 0.5, "Phi-cut gain not available",
                                transform=self.ax_ph.transAxes, ha="center", va="center")

            # Redraw the canvas
            self.canvas.draw()
            self.log_message("Radiation patterns refreshed")
            
        except Exception as e:
            self.log_message(f"Error refreshing patterns: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    def toggle_auto_refresh(self):
        """Toggle auto-refresh of patterns"""
        if self.auto_refresh_var.get():
            self.schedule_auto_refresh()
        else:
            if self.auto_refresh_job:
                self.after_cancel(self.auto_refresh_job)
                self.auto_refresh_job = None

    def schedule_auto_refresh(self):
        """Schedule the next auto-refresh"""
        if self.auto_refresh_var.get():
            self.refresh_patterns_only()
            self.auto_refresh_job = self.after(1500, self.schedule_auto_refresh)

    def apply_sources_from_ui(self):
        """Apply source settings from UI controls"""
        try:
            # Get current excitations
            excitations = self._list_excitations()
            if not excitations:
                self.log_message("No excitations found for source control")
                return
                
            # Prepare magnitude and phase arrays
            magnitudes = []
            phases = []
            
            for i, ex in enumerate(excitations, 1):
                # Get values from UI controls if they exist
                if ex in self.source_controls:
                    mag = self.source_controls[ex]["power"].get()
                    ph = self.source_controls[ex]["phase"].get()
                    magnitudes.append(f"{mag}W")
                    phases.append(f"{ph}deg")
                else:
                    # Use default values
                    magnitudes.append("1W")
                    phases.append("0deg")
            
            # Apply the source settings
            self._edit_sources_with_vars(excitations, magnitudes, phases)
            self.log_message("Source settings applied successfully")
            
            # Refresh patterns to see the changes
            self.refresh_patterns_only()
            
        except Exception as e:
            self.log_message(f"Error applying source settings: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # ------------- Cleanup -------------
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

    def save_parameters(self):
        """Save parameters to JSON file."""
        try:
            all_params = {**self.params, **self.calculated_params}
            with open("antenna_parameters.json", "w") as f:
                json.dump(all_params, f, indent=4)
            self.log_message("Parameters saved to antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error saving parameters: {e}")

    def load_parameters(self):
        """Load parameters from JSON file."""
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
        """Update GUI with loaded parameters."""
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
                        
            # Update calculated parameters display
            self.patches_label.configure(text=f"Number of Patches: {self.calculated_params['num_patches']}")
            self.rows_cols_label.configure(
                text=f"Configuration: {self.calculated_params['rows']} x {self.calculated_params['cols']}"
            )
            self.spacing_label.configure(
                text=f"Spacing: {self.calculated_params['spacing']:.2f} mm ({self.params['spacing_type']})"
            )
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
        """Run the application."""
        try:
            self.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.mainloop()
        finally:
            self.cleanup()


if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()