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
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk

import ansys.aedt.core
from ansys.aedt.core import Desktop, Hfss

# ---------- Appearance ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ModernPatchAntennaDesigner:
    def __init__(self):
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
        self.original_params = {}  # Store original parameters for comparison
        self.optimized = False  # Track if we're in optimized mode
        self.optimization_history = []  # Track optimization steps

        # -------- User Parameters --------
        self.params = {
            "frequency": 10.0,             # GHz
            "gain": 12.0,                  # dBi
            "sweep_start": 8.0,            # GHz
            "sweep_stop": 12.0,            # GHz
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,
            "spacing_type": "lambda/2",
            "substrate_material": "Duroid (tm)",
            "substrate_thickness": 0.5,    # mm
            "metal_thickness": 0.035,      # mm
            "er": 2.2,
            "tan_d": 0.0009,
            "feed_position": "inset",      # edge|inset
            "feed_rel_x": 0.485,
            "probe_radius": 0.40,          # mm (a)
            "coax_er": 1.0,
            "coax_ba_ratio": 2.3,
            "coax_wall_thickness": 0.20,   # mm
            "coax_port_length": 3.0,       # mm
            "antipad_clearance": 0.10,     # mm
            "sweep_type": "Interpolating",  # "Discrete" | "Interpolating" | "Fast"
            "sweep_step": 0.02              # GHz (used in Discrete)
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
        self.setup_gui()

    # ---------------- GUI ----------------
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        self.window.geometry("1600x1000")

        # Configure grid
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        # Header
        header = ctk.CTkFrame(self.window, height=80, fg_color=("gray85", "gray20"))
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
        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Add tabs
        tabs = ["Design Parameters", "Simulation", "Results", "Log", "Parameters Summary"]
        for tab_name in tabs:
            self.tabview.add(tab_name)
            self.tabview.tab(tab_name).grid_columnconfigure(0, weight=1)

        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()
        self.setup_parameters_summary_tab()

        # Status bar
        status = ctk.CTkFrame(self.window, height=40)
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
            text="Note: Simulation time increases with array size.",
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
            text="Simulation Results", 
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=10)

        # Create notebook for different result types
        self.results_notebook = ttk.Notebook(paned_window)
        self.results_notebook.grid(row=1, column=0, sticky="nsew", pady=10)
        
        # Impedance Results Tab
        impedance_frame = ctk.CTkFrame(self.results_notebook, fg_color=("gray92", "gray18"))
        impedance_frame.grid_columnconfigure(0, weight=1)
        impedance_frame.grid_rowconfigure(1, weight=1)
        self.results_notebook.add(impedance_frame, text="Impedance Parameters")
        
        # Radiation Results Tab
        radiation_frame = ctk.CTkFrame(self.results_notebook, fg_color=("gray92", "gray18"))
        radiation_frame.grid_columnconfigure(0, weight=1)
        radiation_frame.grid_rowconfigure(1, weight=1)
        self.results_notebook.add(radiation_frame, text="Radiation Patterns")
        
        # Setup impedance tab
        self.setup_impedance_tab(impedance_frame)
        
        # Setup radiation tab
        self.setup_radiation_tab(radiation_frame)
        
        # Export buttons
        export_frame = ctk.CTkFrame(paned_window)
        export_frame.grid(row=2, column=0, pady=8)
        
        ctk.CTkButton(
            export_frame, 
            text="Export CSV", 
            command=self.export_csv,
            fg_color="#6A5ACD", 
            hover_color="#7B68EE"
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            export_frame, 
            text="Export PNG", 
            command=self.export_png,
            fg_color="#20B2AA", 
            hover_color="#40E0D0"
        ).pack(side="left", padx=8)
        
        # Add optimize button
        ctk.CTkButton(
            export_frame, 
            text="Analyze & Optimize", 
            command=self.analyze_and_optimize,
            fg_color="#FF6347", 
            hover_color="#FF4500"
        ).pack(side="left", padx=8)
        
        # Add reset button
        ctk.CTkButton(
            export_frame, 
            text="Reset to Original", 
            command=self.reset_to_original,
            fg_color="#9370DB", 
            hover_color="#8A2BE2"
        ).pack(side="left", padx=8)

    def setup_impedance_tab(self, parent):
        """Setup the impedance parameters tab"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        
        # Create figure for impedance results
        self.fig_imp = plt.figure(figsize=(10, 8))
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        self.fig_imp.patch.set_facecolor(face)
        
        # Create subplots for S11, VSWR, and impedance
        self.ax_s11 = self.fig_imp.add_subplot(3, 1, 1)
        self.ax_vswr = self.fig_imp.add_subplot(3, 1, 2)
        self.ax_z = self.fig_imp.add_subplot(3, 1, 3)
        
        # Style the axes
        for ax in [self.ax_s11, self.ax_vswr, self.ax_z]:
            ax.set_facecolor(face)
            if ctk.get_appearance_mode() == "Dark":
                ax.tick_params(colors='white')
                ax.xaxis.label.set_color('white')
                ax.yaxis.label.set_color('white')
                ax.title.set_color('white')
                for s in ['bottom', 'top', 'right', 'left']:
                    ax.spines[s].set_color('white')
                ax.grid(color='gray', alpha=0.5)
        
        # Set titles and labels
        self.ax_s11.set_title("S11 Parameter")
        self.ax_s11.set_ylabel("S11 (dB)")
        
        self.ax_vswr.set_title("VSWR")
        self.ax_vswr.set_ylabel("VSWR")
        
        self.ax_z.set_title("Input Impedance")
        self.ax_z.set_xlabel("Frequency (GHz)")
        self.ax_z.set_ylabel("Impedance (Ω)")
        
        # Embed the plot in the GUI
        canvas_frame = ctk.CTkFrame(parent)
        canvas_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        canvas_frame.grid_columnconfigure(0, weight=1)
        canvas_frame.grid_rowconfigure(0, weight=1)
        
        self.canvas_imp = FigureCanvasTkAgg(self.fig_imp, master=canvas_frame)
        self.canvas_imp.get_tk_widget().pack(fill="both", expand=True)
        
        # Add controls for impedance chart
        control_frame = ctk.CTkFrame(parent)
        control_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(control_frame, text="Chart Controls:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        
        # Add frequency range controls
        ctk.CTkLabel(control_frame, text="Freq Range:").pack(side="left", padx=5)
        self.freq_min = ctk.CTkEntry(control_frame, width=80)
        self.freq_min.insert(0, str(self.params["sweep_start"]))
        self.freq_min.pack(side="left", padx=2)
        
        ctk.CTkLabel(control_frame, text="to").pack(side="left", padx=2)
        
        self.freq_max = ctk.CTkEntry(control_frame, width=80)
        self.freq_max.insert(0, str(self.params["sweep_stop"]))
        self.freq_max.pack(side="left", padx=2)
        
        ctk.CTkLabel(control_frame, text="GHz").pack(side="left", padx=2)
        
        ctk.CTkButton(control_frame, text="Update Charts", command=self.update_impedance_charts, 
                     width=120).pack(side="left", padx=10)

    def setup_radiation_tab(self, parent):
        """Setup the radiation patterns tab"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        
        # Create figure for radiation patterns
        self.fig_rad = plt.figure(figsize=(12, 8))
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        self.fig_rad.patch.set_facecolor(face)
        
        # Create subplots for different radiation patterns
        self.ax_2d = self.fig_rad.add_subplot(2, 2, 1, polar=True)
        self.ax_eplane = self.fig_rad.add_subplot(2, 2, 2)
        self.ax_hplane = self.fig_rad.add_subplot(2, 2, 3)
        self.ax_3d = self.fig_rad.add_subplot(2, 2, 4, projection='3d')
        
        # Style the axes
        for ax in [self.ax_2d, self.ax_eplane, self.ax_hplane, self.ax_3d]:
            if hasattr(ax, 'set_facecolor'):
                ax.set_facecolor(face)
            if ctk.get_appearance_mode() == "Dark":
                if hasattr(ax, 'tick_params'):
                    ax.tick_params(colors='white')
                if hasattr(ax, 'xaxis'):
                    ax.xaxis.label.set_color('white')
                if hasattr(ax, 'yaxis'):
                    ax.yaxis.label.set_color('white')
                if hasattr(ax, 'title'):
                    ax.title.set_color('white')
                if hasattr(ax, 'set_zlabel'):
                    ax.zaxis.label.set_color('white')
                for s in ['bottom', 'top', 'right', 'left']:
                    if hasattr(ax, 'spines'):
                        ax.spines[s].set_color('white')
                if hasattr(ax, 'grid'):
                    ax.grid(color='gray', alpha=0.5)
        
        # Set titles
        self.ax_2d.set_title("2D Radiation Pattern (Polar)")
        self.ax_eplane.set_title("E-Plane Cut (YZ Plane)")
        self.ax_eplane.set_xlabel("Angle (degrees)")
        self.ax_eplane.set_ylabel("Gain (dBi)")
        self.ax_hplane.set_title("H-Plane Cut (XY Plane)")
        self.ax_hplane.set_xlabel("Angle (degrees)")
        self.ax_hplane.set_ylabel("Gain (dBi)")
        self.ax_3d.set_title("3D Radiation Pattern")
        
        # Embed the plot in the GUI
        canvas_frame = ctk.CTkFrame(parent)
        canvas_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        canvas_frame.grid_columnconfigure(0, weight=1)
        canvas_frame.grid_rowconfigure(0, weight=1)
        
        self.canvas_rad = FigureCanvasTkAgg(self.fig_rad, master=canvas_frame)
        self.canvas_rad.get_tk_widget().pack(fill="both", expand=True)
        
        # Add controls for radiation charts
        control_frame = ctk.CTkFrame(parent)
        control_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(control_frame, text="Radiation Controls:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        
        # Frequency selection
        ctk.CTkLabel(control_frame, text="Frequency:").pack(side="left", padx=5)
        self.rad_freq = ctk.CTkEntry(control_frame, width=80)
        self.rad_freq.insert(0, str(self.params["frequency"]))
        self.rad_freq.pack(side="left", padx=2)
        ctk.CTkLabel(control_frame, text="GHz").pack(side="left", padx=2)
        
        ctk.CTkButton(control_frame, text="Update Patterns", command=self.update_radiation_patterns, 
                     width=120).pack(side="left", padx=10)

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

    def setup_parameters_summary_tab(self):
        tab = self.tabview.tab("Parameters Summary")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            main, 
            text="Antenna Parameters Summary", 
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=10)

        # Create a frame for the parameter treeview
        tree_frame = ctk.CTkFrame(main)
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        # Create a treeview to display parameters
        self.param_tree = ttk.Treeview(tree_frame, columns=("Value", "Unit"), show="tree", height=20)
        self.param_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.param_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.param_tree.configure(yscrollcommand=scrollbar.set)

        # Configure treeview style
        style = ttk.Style()
        style.configure("Treeview", 
                        background="#2b2b2b" if ctk.get_appearance_mode() == "Dark" else "#ffffff",
                        fieldbackground="#2b2b2b" if ctk.get_appearance_mode() == "Dark" else "#ffffff",
                        foreground="white" if ctk.get_appearance_mode() == "Dark" else "black")
        
        style.configure("Treeview.Heading", 
                        background="#3b3b3b" if ctk.get_appearance_mode() == "Dark" else "#e0e0e0",
                        foreground="white" if ctk.get_appearance_mode() == "Dark" else "black")

        # Set column headings
        self.param_tree.heading("#0", text="Parameter")
        self.param_tree.heading("Value", text="Value")
        self.param_tree.heading("Unit", text="Unit")

        # Set column widths
        self.param_tree.column("#0", width=250, minwidth=200)
        self.param_tree.column("Value", width=150, minwidth=100)
        self.param_tree.column("Unit", width=80, minwidth=60)

        # Populate the treeview with parameters
        self.update_parameters_summary()

        # Button to refresh the summary
        btn_frame = ctk.CTkFrame(main)
        btn_frame.grid(row=2, column=0, pady=8)
        
        ctk.CTkButton(
            btn_frame, 
            text="Refresh Summary", 
            command=self.update_parameters_summary
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            btn_frame, 
            text="Export to CSV", 
            command=self.export_parameters_csv
        ).pack(side="left", padx=8)

    def update_parameters_summary(self):
        """Update the parameters summary treeview"""
        # Clear existing items
        for item in self.param_tree.get_children():
            self.param_tree.delete(item)
        
        # Add user parameters
        user_params = self.param_tree.insert("", "end", text="User Parameters", values=("", ""))
        
        for key, value in self.params.items():
            if key not in ["non_graphical", "aedt_version"]:  # Skip some parameters
                unit = self.get_parameter_unit(key)
                self.param_tree.insert(user_params, "end", text=key.replace("_", " ").title(), 
                                      values=(f"{value}", unit))
        
        # Add calculated parameters
        calc_params = self.param_tree.insert("", "end", text="Calculated Parameters", values=("", ""))
        
        for key, value in self.calculated_params.items():
            unit = self.get_parameter_unit(key)
            self.param_tree.insert(calc_params, "end", text=key.replace("_", " ").title(), 
                                  values=(f"{value:.4f}" if isinstance(value, float) else f"{value}", unit))
        
        # Expand all items
        self.param_tree.item(user_params, open=True)
        self.param_tree.item(calc_params, open=True)

    def get_parameter_unit(self, param_name):
        """Get the unit for a parameter"""
        units = {
            "frequency": "GHz",
            "gain": "dBi",
            "sweep_start": "GHz",
            "sweep_stop": "GHz",
            "substrate_thickness": "mm",
            "metal_thickness": "mm",
            "er": "",
            "tan_d": "",
            "probe_radius": "mm",
            "coax_wall_thickness": "mm",
            "coax_port_length": "mm",
            "antipad_clearance": "mm",
            "sweep_step": "GHz",
            "patch_length": "mm",
            "patch_width": "mm",
            "spacing": "mm",
            "lambda_g": "mm",
            "feed_offset": "mm",
            "substrate_width": "mm",
            "substrate_length": "mm",
            "num_patches": "",
            "rows": "",
            "cols": ""
        }
        
        return units.get(param_name, "")

    def export_parameters_csv(self):
        """Export parameters to CSV file"""
        try:
            with open("antenna_parameters_summary.csv", "w") as f:
                f.write("Category,Parameter,Value,Unit\n")
                
                # Write user parameters
                for key, value in self.params.items():
                    if key not in ["non_graphical", "aedt_version"]:
                        unit = self.get_parameter_unit(key)
                        f.write(f"User,{key},{value},{unit}\n")
                
                # Write calculated parameters
                for key, value in self.calculated_params.items():
                    unit = self.get_parameter_unit(key)
                    f.write(f"Calculated,{key},{value},{unit}\n")
            
            self.log_message("Parameters summary exported to antenna_parameters_summary.csv")
            messagebox.showinfo("Export Successful", "Parameters summary exported to CSV file.")
            
        except Exception as e:
            self.log_message(f"Error exporting parameters to CSV: {e}")
            messagebox.showerror("Export Error", f"Failed to export parameters: {e}")

    def update_impedance_charts(self):
        """Update impedance charts with new frequency range"""
        try:
            # Get new frequency range
            freq_min = float(self.freq_min.get())
            freq_max = float(self.freq_max.get())
            
            # Update sweep parameters
            self.params["sweep_start"] = freq_min
            self.params["sweep_stop"] = freq_max
            
            # Re-run simulation or update charts if data is available
            if self.simulation_data is not None:
                self.plot_impedance_results()
                
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numeric values for frequency range.")
        except Exception as e:
            self.log_message(f"Error updating impedance charts: {e}")

    def update_radiation_patterns(self):
        """Update radiation patterns with new frequency"""
        try:
            # Get new frequency
            freq = float(self.rad_freq.get())
            
            # Update plots if data is available
            if hasattr(self, 'original_theta_data') and hasattr(self, 'original_phi_data'):
                self.plot_radiation_patterns(freq)
                
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid numeric value for frequency.")
        except Exception as e:
            self.log_message(f"Error updating radiation patterns: {e}")

    def plot_impedance_results(self):
        """Plot impedance results (S11, VSWR, impedance)"""
        try:
            # Clear previous plots
            self.ax_s11.clear()
            self.ax_vswr.clear()
            self.ax_z.clear()
            
            # Plot S11
            if self.simulation_data is not None:
                freqs = self.simulation_data[:, 0]
                s11 = self.simulation_data[:, 1]
                
                self.ax_s11.plot(freqs, s11, 'b-', linewidth=2, label='S11')
                self.ax_s11.axhline(y=-10, color='r', linestyle='--', alpha=0.7, label='-10 dB')
                self.ax_s11.set_xlabel("Frequency (GHz)")
                self.ax_s11.set_ylabel("S11 (dB)")
                self.ax_s11.set_title("S11 Parameter")
                self.ax_s11.legend()
                self.ax_s11.grid(True, alpha=0.5)
                
                # Mark center frequency
                cf = float(self.params["frequency"])
                self.ax_s11.axvline(x=cf, color='g', linestyle='--', alpha=0.7)
                self.ax_s11.text(cf, self.ax_s11.get_ylim()[1]*0.9, f'{cf} GHz', color='g')
                
                # Calculate and plot VSWR
                vswr = (1 + 10**(s11/20)) / (1 - 10**(s11/20))
                self.ax_vswr.plot(freqs, vswr, 'g-', linewidth=2, label='VSWR')
                self.ax_vswr.axhline(y=2, color='r', linestyle='--', alpha=0.7, label='VSWR=2')
                self.ax_vswr.set_xlabel("Frequency (GHz)")
                self.ax_vswr.set_ylabel("VSWR")
                self.ax_vswr.set_title("Voltage Standing Wave Ratio")
                self.ax_vswr.legend()
                self.ax_vswr.grid(True, alpha=0.5)
                self.ax_vswr.axvline(x=cf, color='g', linestyle='--', alpha=0.7)
            
            # Update canvas
            self.fig_imp.tight_layout()
            self.canvas_imp.draw()
            
        except Exception as e:
            self.log_message(f"Error plotting impedance results: {e}")

    def plot_radiation_patterns(self, frequency):
        """Plot radiation patterns at specified frequency"""
        try:
            # Clear previous plots
            self.ax_2d.clear()
            self.ax_eplane.clear()
            self.ax_hplane.clear()
            self.ax_3d.clear()
            
            # Get radiation pattern data for E-plane (YZ plane, phi=0°)
            eplane_theta, eplane_gain = self._get_gain_cut(frequency, cut="theta", fixed_angle_deg=0.0)
            
            # Get radiation pattern data for H-plane (XY plane, theta=90°)
            hplane_phi, hplane_gain = self._get_gain_cut(frequency, cut="phi", fixed_angle_deg=90.0)
            
            if eplane_theta is not None and eplane_gain is not None:
                # Convert to radians for polar plot
                theta_rad = np.radians(eplane_theta)
                
                # Plot 2D radiation pattern (polar)
                self.ax_2d.plot(theta_rad, eplane_gain, 'b-', linewidth=2)
                self.ax_2d.set_title("2D Radiation Pattern (Polar)")
                self.ax_2d.grid(True)
                
                # Plot E-plane cut (YZ plane)
                self.ax_eplane.plot(eplane_theta, eplane_gain, 'b-', linewidth=2, label='E-plane (YZ)')
                self.ax_eplane.set_xlabel("Angle (degrees)")
                self.ax_eplane.set_ylabel("Gain (dBi)")
                self.ax_eplane.set_title("E-Plane Cut (YZ Plane)")
                self.ax_eplane.legend()
                self.ax_eplane.grid(True, alpha=0.5)
            
            if hplane_phi is not None and hplane_gain is not None:
                # Plot H-plane cut (XY plane)
                self.ax_hplane.plot(hplane_phi, hplane_gain, 'r-', linewidth=2, label='H-plane (XY)')
                self.ax_hplane.set_xlabel("Angle (degrees)")
                self.ax_hplane.set_ylabel("Gain (dBi)")
                self.ax_hplane.set_title("H-Plane Cut (XY Plane)")
                self.ax_hplane.legend()
                self.ax_hplane.grid(True, alpha=0.5)
            
            # Create a 3D radiation pattern visualization
            # This is a simplified representation using the available data
            if eplane_theta is not None and eplane_gain is not None and hplane_phi is not None and hplane_gain is not None:
                # Create a meshgrid for 3D plotting
                theta = np.linspace(0, 2*np.pi, 36)
                phi = np.linspace(0, np.pi, 18)
                theta, phi = np.meshgrid(theta, phi)
                
                # Interpolate gain values to create a 3D pattern
                # This is a simplified approximation
                max_gain = max(np.max(eplane_gain) if eplane_gain is not None else 0, 
                              np.max(hplane_gain) if hplane_gain is not None else 0)
                
                # Create a radiation pattern shape based on the E and H plane patterns
                r = 5 * (1 + 0.5 * np.sin(phi) * np.cos(2*theta))  # Example pattern shape
                
                # Scale by gain (simplified)
                r = r * (max_gain / 10 if max_gain > 0 else 1)
                
                # Convert to Cartesian coordinates
                x = r * np.sin(phi) * np.cos(theta)
                y = r * np.sin(phi) * np.sin(theta)
                z = r * np.cos(phi)
                
                # Plot 3D radiation pattern
                self.ax_3d.plot_surface(x, y, z, cmap='viridis', alpha=0.8, edgecolor='none')
                self.ax_3d.set_title("3D Radiation Pattern")
                self.ax_3d.set_xlabel("X")
                self.ax_3d.set_ylabel("Y")
                self.ax_3d.set_zlabel("Z")
            
            # Update canvas
            self.fig_rad.tight_layout()
            self.canvas_rad.draw()
            
        except Exception as e:
            self.log_message(f"Error plotting radiation patterns: {e}")
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
            if hasattr(self, 'fig_imp'):
                self.fig_imp.savefig("impedance_results.png", dpi=300, bbox_inches='tight')
                self.log_message("Impedance plot saved to impedance_results.png")
            
            if hasattr(self, 'fig_rad'):
                self.fig_rad.savefig("radiation_patterns.png", dpi=300, bbox_inches='tight')
                self.log_message("Radiation plot saved to radiation_patterns.png")
                
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
                             "sweep_step"]:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
                elif key in ["spacing_type", "substrate_material", "feed_position", "sweep_type"]:
                    self.params[key] = widget.get()
                else:
                    if isinstance(widget, ctk