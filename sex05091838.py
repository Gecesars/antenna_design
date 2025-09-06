# -*- coding: utf-8 -*-
import os
import tempfile
from datetime import datetime
import math
import json
import traceback
import queue
import threading
from typing import Tuple, List
import tkinter as tk
from tkinter import messagebox

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk

import ansys.aedt.core
from ansys.aedt.core import Desktop, Hfss

# ---------- Aparência Moderna ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ModernPatchAntennaDesigner:
    def __init__(self):
        self.hfss = None
        self.desktop: Desktop | None = None
        self.temp_folder = None
        self.project_path = ""
        self.project_display_name = "patch_array"
        self.design_base_name = "patch_array"
        self.log_queue = queue.Queue()
        self.is_simulation_running = False
        self.save_project = False
        self.stop_simulation = False
        self.simulated_resonance_freq = None
        self.simulation_data = None

        # -------- Parâmetros do usuário --------
        self.params = {
            "frequency": 10.0,             # GHz
            "gain": 12.0,                  # dBi
            "sweep_start": 8.0,            # GHz
            "sweep_stop": 12.0,            # GHz
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,
            "spacing_type": "lambda/2",
            "substrate_material": "Duroid (tm)",   # nome que existe
            "substrate_thickness": 0.5,    # mm
            "metal_thickness": 0.035,      # mm
            "er": 2.2,
            "tan_d": 0.0009,
            "feed_position": "inset",      # edge|inset
            "feed_rel_x": 0.485,           # fração*W a partir da borda esquerda
            "probe_radius": 0.40,          # mm (a)
            "coax_er": 1.0,                # só para referência
            "coax_ba_ratio": 2.3,          # b/a
            "coax_wall_thickness": 0.20,   # mm
            "coax_port_length": 3.0,       # mm  (Lp)
            "antipad_clearance": 0.10      # mm
        }

        # -------- Parâmetros calculados --------
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

        # -------- Pesos de excitação (GUI) --------
        self.port_params: List[dict] = [{"amp": 1.0, "phase": 0.0} for _ in range(self.calculated_params["num_patches"])]
        self.port_entries: List[tuple] = []

        self.c = 299792458.0
        self.setup_gui()

    # ---------------- GUI Modernizada ----------------
    def _maximize_with_taskbar(self):
        """Maximiza mantendo a barra de tarefas visível."""
        try:
            self.window.attributes("-fullscreen", False)
        except Exception:
            pass
        try:
            self.window.state("zoomed")   # Windows
            return
        except Exception:
            pass
        try:
            self.window.attributes("-zoomed", True)  # Linux/GNOME
            return
        except Exception:
            pass
        try:
            w = self.window.winfo_screenwidth()
            h = self.window.winfo_screenheight()
            self.window.geometry(f"{w}x{h}+0+0")
        except Exception:
            pass

    def setup_gui(self):
        # Configuração da janela principal
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        self.window.geometry("1600x1000")
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Configuração do layout principal
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)
        
        # Barra de título moderna
        title_bar = ctk.CTkFrame(self.window, height=70, fg_color=("#1E3F66", "#1A3560"), corner_radius=0)
        title_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        title_bar.grid_propagate(False)
        title_bar.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(title_bar, text="Patch Antenna Array Designer",
                     font=ctk.CTkFont(size=24, weight="bold", family="Segoe UI"),
                     text_color="white").pack(expand=True, pady=20)
        
        # Área principal avec tabs
        self.tabview = ctk.CTkTabview(self.window, fg_color=("gray90", "gray15"), 
                                     segmented_button_fg_color=("#3B8ED0", "#1F6AA5"),
                                     segmented_button_selected_color=("#1F6AA5", "#1F538D"),
                                     segmented_button_selected_hover_color=("#1F538D", "#1A4A7F"),
                                     text_color=("white", "white"))
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Adicionar abas
        tabs = ["Design Parameters", "Simulation", "Results", "Log"]
        for name in tabs:
            self.tabview.add(name)
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)
        
        # Configurar cada aba
        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()
        
        # Barra de status moderna
        status_bar = ctk.CTkFrame(self.window, height=40, fg_color=("#1E3F66", "#1A3560"), corner_radius=0)
        status_bar.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        status_bar.grid_propagate(False)
        status_bar.grid_columnconfigure(0, weight=1)
        
        self.status_label = ctk.CTkLabel(status_bar, text="Ready to calculate parameters",
                                         font=ctk.CTkFont(weight="bold", family="Segoe UI"),
                                         text_color="white")
        self.status_label.pack(side="left", padx=15, pady=8)
        
        # Iniciar processamento da fila de log
        self.process_log_queue()
        self._maximize_with_taskbar()

    def create_section(self, parent, title, row, column, padx=10, pady=10, colspan=1):
        section = ctk.CTkFrame(parent, fg_color=("gray95", "gray20"), corner_radius=8)
        section.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady, columnspan=colspan)
        section.grid_columnconfigure(0, weight=1)
        
        # Título da seção com estilo moderno
        title_frame = ctk.CTkFrame(section, fg_color=("gray90", "gray25"), height=40, corner_radius=6)
        title_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 0))
        title_frame.grid_propagate(False)
        
        ctk.CTkLabel(title_frame, text=title,
                     font=ctk.CTkFont(size=14, weight="bold", family="Segoe UI"),
                     text_color=("gray20", "gray90")).pack(expand=True)
        
        return section

    def setup_parameters_tab(self):
        tab = self.tabview.tab("Design Parameters")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame principal com scroll
        main = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main.grid_columnconfigure((0, 1), weight=1, uniform="col")
        main.grid_rowconfigure(0, weight=1)

        # ---------------- Antenna Parameters ----------------
        sec_ant = self.create_section(main, "Antenna Parameters", 0, 0, pady=12)
        entries = []
        r = 1

        def add(section, label, key, value, row, combo=None, check=False, tooltip=None):
            frame = ctk.CTkFrame(section, fg_color="transparent")
            frame.grid(row=row, column=0, sticky="ew", padx=10, pady=6)
            frame.grid_columnconfigure(0, weight=1)
            frame.grid_columnconfigure(1, weight=2)
            
            ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(weight="bold", size=12),
                         anchor="w").grid(row=0, column=0, padx=(0, 10), pady=2, sticky="w")
            
            if combo:
                var = ctk.StringVar(value=value)
                w = ctk.CTkComboBox(frame, values=combo, variable=var, width=220, 
                                   dropdown_fg_color=("gray90", "gray25"),
                                   button_color=("#3B8ED0", "#1F6AA5"))
                w.grid(row=0, column=1, padx=(0, 10), pady=2, sticky="ew")
                entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                w = ctk.CTkCheckBox(frame, text="", variable=var, 
                                   checkbox_width=18, checkbox_height=18,
                                   fg_color=("#3B8ED0", "#1F6AA5"), hover_color=("#1F6AA5", "#1F538D"))
                w.grid(row=0, column=1, padx=(0, 10), pady=2, sticky="w")
                entries.append((key, var))
            else:
                w = ctk.CTkEntry(frame, width=220, fg_color=("white", "gray25"),
                                border_color=("#3B8ED0", "#1F6AA5"))
                w.insert(0, str(value))
                w.grid(row=0, column=1, padx=(0, 10), pady=2, sticky="ew")
                entries.append((key, w))
            
            if tooltip:
                # Adicionar ícone de informação com tooltip (implementação simplificada)
                info_btn = ctk.CTkButton(frame, text="?", width=20, height=20,
                                        fg_color=("gray70", "gray40"), hover_color=("gray60", "gray50"),
                                        font=ctk.CTkFont(size=10, weight="bold"))
                info_btn.grid(row=0, column=2, padx=(0, 5), pady=2)
            
            return row + 1

        r = add(sec_ant, "Central Frequency (GHz):", "frequency", self.params["frequency"], r, tooltip="Operating frequency in GHz")
        r = add(sec_ant, "Desired Gain (dBi):", "gain", self.params["gain"], r)
        r = add(sec_ant, "Sweep Start (GHz):", "sweep_start", self.params["sweep_start"], r)
        r = add(sec_ant, "Sweep Stop (GHz):", "sweep_stop", self.params["sweep_stop"], r)
        r = add(sec_ant, "Patch Spacing:", "spacing_type", self.params["spacing_type"], r,
                combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        # ---------------- Substrate Parameters ----------------
        sec_sub = self.create_section(main, "Substrate Parameters", 0, 1, pady=12)
        r = 1
        r = add(sec_sub, "Substrate Material:", "substrate_material",
                self.params["substrate_material"], r,
                combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        r = add(sec_sub, "Relative Permittivity (εr):", "er", self.params["er"], r)
        r = add(sec_sub, "Loss Tangent (tan δ):", "tan_d", self.params["tan_d"], r)
        r = add(sec_sub, "Substrate Thickness (mm):", "substrate_thickness", self.params["substrate_thickness"], r)
        r = add(sec_sub, "Metal Thickness (mm):", "metal_thickness", self.params["metal_thickness"], r)

        # ---------------- Coaxial Feed Parameters ----------------
        sec_coax = self.create_section(main, "Coaxial Feed Parameters", 1, 0, pady=12)
        r = 1
        r = add(sec_coax, "Feed position type:", "feed_position", self.params["feed_position"], r,
                combo=["inset", "edge"])
        r = add(sec_coax, "Feed relative X (0..1):", "feed_rel_x", self.params["feed_rel_x"], r)
        r = add(sec_coax, "Inner radius a (mm):", "probe_radius", self.params["probe_radius"], r)
        r = add(sec_coax, "b/a ratio:", "coax_ba_ratio", self.params["coax_ba_ratio"], r)
        r = add(sec_coax, "Shield wall (mm):", "coax_wall_thickness", self.params["coax_wall_thickness"], r)
        r = add(sec_coax, "Port length below GND Lp (mm):", "coax_port_length", self.params["coax_port_length"], r)
        r = add(sec_coax, "Anti-pad clearance (mm):", "antipad_clearance", self.params["antipad_clearance"], r)

        # ---------------- Simulation Settings ----------------
        sec_sim = self.create_section(main, "Simulation Settings", 1, 1, pady=12)
        r = 1
        r = add(sec_sim, "CPU Cores:", "cores", self.params["cores"], r)
        r = add(sec_sim, "Show HFSS Interface:", "show_gui", not self.params["non_graphical"], r, check=True)
        r = add(sec_sim, "Save Project:", "save_project", self.save_project, r, check=True)

        self.entries = entries

        # ---------------- Calculated Parameters ----------------
        sec_calc = self.create_section(main, "Calculated Parameters", 2, 0, pady=12, colspan=2)
        grid = ctk.CTkFrame(sec_calc, fg_color="transparent")
        grid.grid(row=1, column=0, sticky="nsew", padx=15, pady=10)
        grid.columnconfigure((0, 1, 2, 3), weight=1, uniform="col")
        
        # Organização em grid 2x4 para melhor visualização
        self.patches_label = ctk.CTkLabel(grid, text="Number of Patches: 4", 
                                         font=ctk.CTkFont(weight="bold", size=12))
        self.patches_label.grid(row=0, column=0, sticky="w", pady=4, padx=10)
        
        self.rows_cols_label = ctk.CTkLabel(grid, text="Configuration: 2 x 2", 
                                           font=ctk.CTkFont(weight="bold", size=12))
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=4, padx=10)
        
        self.spacing_label = ctk.CTkLabel(grid, text="Spacing: -- mm", 
                                         font=ctk.CTkFont(weight="bold", size=12))
        self.spacing_label.grid(row=0, column=2, sticky="w", pady=4, padx=10)
        
        self.dimensions_label = ctk.CTkLabel(grid, text="Patch Dimensions: -- x -- mm",
                                            font=ctk.CTkFont(weight="bold", size=12))
        self.dimensions_label.grid(row=0, column=3, sticky="w", pady=4, padx=10)
        
        self.lambda_label = ctk.CTkLabel(grid, text="Guided Wavelength: -- mm", 
                                        font=ctk.CTkFont(weight="bold", size=12))
        self.lambda_label.grid(row=1, column=0, sticky="w", pady=4, padx=10)
        
        self.feed_offset_label = ctk.CTkLabel(grid, text="Feed Offset (y): -- mm", 
                                             font=ctk.CTkFont(weight="bold", size=12))
        self.feed_offset_label.grid(row=1, column=1, sticky="w", pady=4, padx=10)
        
        self.substrate_dims_label = ctk.CTkLabel(grid, text="Substrate Dimensions: -- x -- mm",
                                                font=ctk.CTkFont(weight="bold", size=12))
        self.substrate_dims_label.grid(row=1, column=2, sticky="w", pady=4, padx=10, columnspan=2)

        # Botões de ação
        btns = ctk.CTkFrame(sec_calc, fg_color="transparent")
        btns.grid(row=2, column=0, sticky="ew", padx=15, pady=12, columnspan=2)
        btns.grid_columnconfigure((0, 1, 2), weight=1)
        
        ctk.CTkButton(btns, text="Calculate Parameters", command=self.calculate_parameters,
                      fg_color=("#2E8B57", "#2E8B57"), hover_color=("#3CB371", "#3CB371"), 
                      height=36, font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=8, pady=5)
        
        ctk.CTkButton(btns, text="Save Parameters", command=self.save_parameters,
                      fg_color=("#4169E1", "#4169E1"), hover_color=("#6495ED", "#6495ED"),
                      height=36, font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=8, pady=5)
        
        ctk.CTkButton(btns, text="Load Parameters", command=self.load_parameters,
                      fg_color=("#FF8C00", "#FF8C00"), hover_color=("#FFA500", "#FFA500"),
                      height=36, font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=8, pady=5)

        # ---------------- Post-Processing Weights ----------------
        self._build_postproc_section(main, row=3)

    def _build_postproc_section(self, parent, row: int):
        sec_pp = self.create_section(parent, "Post-Processing: Port Power & Phase", row, 0, colspan=2)

        # Cabeçalho da tabela
        header = ctk.CTkFrame(sec_pp, fg_color=("gray88", "gray22"), height=40)
        header.grid(row=1, column=0, sticky="ew", padx=12, pady=(8, 0))
        header.grid_columnconfigure((0, 1, 2), weight=1)
        
        for i, t in enumerate(["Port #", "Power (rel.)", "Phase (deg)"]):
            ctk.CTkLabel(header, text=t, font=ctk.CTkFont(weight="bold", size=12)).grid(
                row=0, column=i, padx=10, pady=10, sticky="w" if i==0 else "ew")
        
        # Tabela de portas com scroll
        self.ports_table = ctk.CTkScrollableFrame(sec_pp, height=160, fg_color=("gray95", "gray18"))
        self.ports_table.grid(row=2, column=0, sticky="nsew", padx=12, pady=8)
        self.ports_table.grid_columnconfigure((0, 1, 2), weight=1)

        # Botões de ação
        row_btn = ctk.CTkFrame(sec_pp, fg_color="transparent")
        row_btn.grid(row=3, column=0, sticky="ew", padx=12, pady=(2, 10))
        row_btn.grid_columnconfigure((0, 1), weight=1)
        
        ctk.CTkButton(row_btn, text="Reset (Broadside)", command=self._reset_port_weights,
                      fg_color=("gray70", "gray40"), hover_color=("gray60", "gray50"),
                      height=32).grid(row=0, column=0, padx=6, sticky="w")
        
        ctk.CTkButton(row_btn, text="Apply to HFSS Vars", command=self.apply_postproc_variables,
                      fg_color=("#6A5ACD", "#6A5ACD"), hover_color=("#7B68EE", "#7B68EE"),
                      height=32).grid(row=0, column=1, padx=6, sticky="e")

        self.update_ports_table(self.calculated_params["num_patches"])

    def update_ports_table(self, n_ports: int):
        self._sync_port_params_from_ui()
        cur = len(self.port_params)
        if n_ports > cur:
            self.port_params += [{"amp": 1.0, "phase": 0.0} for _ in range(n_ports - cur)]
        elif n_ports < cur:
            self.port_params = self.port_params[:n_ports]
        
        # Limpar tabela atual
        for widget in self.ports_table.winfo_children():
            widget.destroy()
            
        self.port_entries.clear()
        
        # Criar nova tabela
        for i in range(n_ports):
            idx = i + 1
            
            # Número da porta
            port_label = ctk.CTkLabel(self.ports_table, text=f"P{idx}", 
                                     font=ctk.CTkFont(weight="bold"))
            port_label.grid(row=i, column=0, padx=10, pady=4, sticky="w")
            
            # Campo de amplitude
            e_amp = ctk.CTkEntry(self.ports_table, width=100, 
                                fg_color=("white", "gray25"),
                                border_color=("#3B8ED0", "#1F6AA5"))
            e_amp.insert(0, str(self.port_params[i]["amp"]))
            e_amp.grid(row=i, column=1, padx=10, pady=4, sticky="ew")
            
            # Campo de fase
            e_ph = ctk.CTkEntry(self.ports_table, width=100,
                               fg_color=("white", "gray25"),
                               border_color=("#3B8ED0", "#1F6AA5"))
            e_ph.insert(0, str(self.port_params[i]["phase"]))
            e_ph.grid(row=i, column=2, padx=10, pady=4, sticky="ew")
            
            self.port_entries.append((e_amp, e_ph))
            
        # Configurar pesos das colunas
        self.ports_table.grid_columnconfigure(0, weight=1)
        self.ports_table.grid_columnconfigure(1, weight=2)
        self.ports_table.grid_columnconfigure(2, weight=2)

    def _reset_port_weights(self):
        for p in self.port_params:
            p["amp"] = 1.0
            p["phase"] = 0.0
        self.update_ports_table(len(self.port_params))
        self.log_message("Post-processing weights reset to broadside (amp=1, phase=0).")

    def _sync_port_params_from_ui(self):
        if not self.port_entries:
            return
        for i, (e_amp, e_ph) in enumerate(self.port_entries):
            try:
                amp = float(e_amp.get())
            except Exception:
                amp = 1.0
            try:
                ph = float(e_ph.get())
            except Exception:
                ph = 0.0
            if i < len(self.port_params):
                self.port_params[i]["amp"] = amp
                self.port_params[i]["phase"] = ph

    def apply_postproc_variables(self):
        """Botão da GUI – grava Pow_Pi e Phi_Pi conforme tabela atual."""
        try:
            self._sync_port_params_from_ui()
            n = len(self.port_params)
            if not self.hfss:
                self.log_message("HFSS project not available yet. Run a simulation first or open a project.")
                return
            for i in range(n):
                idx = i + 1
                amp = self.port_params[i]["amp"]
                ph = self.port_params[i]["phase"]
                self.hfss[f"Pow_P{idx}"] = f"{amp}W"
                self.hfss[f"Phi_P{idx}"] = f"{ph}deg"
            self.log_message(f"Applied post-processing vars for {n} ports (Pow_Pi=1W default, Phi_Pi=0deg).")
        except Exception as e:
            self.log_message(f"Error applying post-processing variables: {e}")

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
                             "probe_radius", "coax_ba_ratio", "coax_wall_thickness",
                             "coax_port_length", "antipad_clearance", "feed_rel_x"]:
                    if isinstance(widget, ctk.CTkEntry):
                        self.params[key] = float(widget.get())
                elif key in ["spacing_type", "substrate_material", "feed_position"]:
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
        f = frequency_ghz * 1e9
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"]) / 1000.0
        W = self.c / (2 * f) * math.sqrt(2 / (er + 1))
        eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W) ** -0.5
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        L_eff = self.c / (2 * f * math.sqrt(eeff))
        L = L_eff - 2 * dL
        lambda_g = self.c / (f * math.sqrt(eeff))
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)

    def _size_array_from_gain(self):
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
        return rows, cols, N_req

    def calculate_substrate_size(self):
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
            self.log_message(f"Array sizing -> target gain {self.params['gain']} dBi, N_req≈{N_req}, "
                             f"layout {rows}x{cols} (= {rows*cols} patches)")
            self.calculated_params["feed_offset"] = 0.30 * L_mm
            self.calculate_substrate_size()
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
            self.update_ports_table(self.calculated_params["num_patches"])
        except Exception as e:
            self.status_label.configure(text=f"Error in calculation: {e}")
            self.log_message(f"Error in calculation: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    # --------- AEDT helpers ---------
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

    def _open_or_create_project(self):
        if self.desktop is None:
            self.desktop = Desktop(version=self.params["aedt_version"],
                                   non_graphical=self.params["non_graphical"],
                                   new_desktop=True)
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
        if self.project_display_name in open_names:
            idx = open_names.index(self.project_display_name)
            proj_obj = open_objs[idx]
            new_design = self.design_base_name
            try:
                tmp = Hfss(project=proj_obj, non_graphical=self.params["non_graphical"],
                           version=self.params["aedt_version"])
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
            self.hfss = Hfss(project=proj_obj, design=new_design, solution_type="DrivenModal",
                             version=self.params["aedt_version"], non_graphical=self.params["non_graphical"])
            try:
                self.project_path = proj_obj.GetPath()
            except Exception:
                self.project_path = ""
            self.log_message(f"Using existing project '{self.project_display_name}', created design '{new_design}'")
            return
        if self.temp_folder is None:
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.project_path = os.path.join(self.temp_folder.name, f"{self.project_display_name}_{ts}.aedt")
        self.hfss = Hfss(project=self.project_path, design=self.design_base_name, solution_type="DrivenModal",
                         version=self.params["aedt_version"], non_graphical=self.params["non_graphical"])
        self.log_message(f"Created new project: {self.project_path} (design '{self.design_base_name}')")

    # ---------- (RESTAURADO) set de variáveis de projeto ----------
    def _set_design_variables(self, L, W, spacing, rows, cols, h_sub, sub_w, sub_l):
        a = float(self.params["probe_radius"])
        ba = float(self.params["coax_ba_ratio"])
        b = a * ba
        wall = float(self.params["coax_wall_thickness"])
        Lp = float(self.params["coax_port_length"])
        clear = float(self.params["antipad_clearance"])

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
        # opcional: fator de correção de indutância do pino
        self.hfss["probeK"] = "0.3"

        self.log_message(f"Air coax set: a={a:.3f} mm, b={b:.3f} mm (b/a={ba:.3f}≈2.3 → ~50 Ω)")
        return a, b, wall, Lp, clear

    # ---------- Coax + Porto Lumped ----------
    def _create_coax_feed_lumped(self, ground, substrate, x_feed: float, y_feed: float,
                                 name_prefix: str):
        """Feed coaxial em ar com porta lumped anelar (linha radial totalmente dentro)."""
        try:
            a_val = float(self.params["probe_radius"])
            b_val = a_val * float(self.params["coax_ba_ratio"])
            wall_val = float(self.params["coax_wall_thickness"])
            Lp_val = float(self.params["coax_port_length"])
            h_sub_val = float(self.params["substrate_thickness"])
            clear_val = float(self.params["antipad_clearance"])
            if b_val - a_val < 0.02:
                b_val = a_val + 0.02

            # PINO condutor
            pin = self.hfss.modeler.create_cylinder(
                orientation="Z", origin=[x_feed, y_feed, -Lp_val],
                radius=a_val, height=h_sub_val + Lp_val + 0.001,
                name=f"{name_prefix}_Pin", material="copper")

            # Blindagem (tubo)
            shield_outer = self.hfss.modeler.create_cylinder(
                orientation="Z", origin=[x_feed, y_feed, -Lp_val],
                radius=b_val + wall_val, height=Lp_val,
                name=f"{name_prefix}_ShieldOuter", material="copper")
            shield_inner_void = self.hfss.modeler.create_cylinder(
                orientation="Z", origin=[x_feed, y_feed, -Lp_val],
                radius=b_val, height=Lp_val,
                name=f"{name_prefix}_ShieldInnerVoid", material="vacuum")
            self.hfss.modeler.subtract(shield_outer, [shield_inner_void], keep_originals=False)
            shield = shield_outer

            # furo no substrato + anti-pad no GND
            hole_r = b_val + clear_val
            sub_hole = self.hfss.modeler.create_cylinder(
                orientation="Z", origin=[x_feed, y_feed, 0.0],
                radius=hole_r, height=h_sub_val, name=f"{name_prefix}_SubHole", material="vacuum")
            self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=False)
            g_hole = self.hfss.modeler.create_circle(
                orientation="XY", origin=[x_feed, y_feed, 0.0], radius=hole_r,
                name=f"{name_prefix}_GndHole", material="vacuum")
            self.hfss.modeler.subtract(ground, [g_hole], keep_originals=False)

            # folja do porto (anel a..b) em z = -Lp
            port_ring = self.hfss.modeler.create_circle(
                orientation="XY", origin=[x_feed, y_feed, -Lp_val],
                radius=b_val, name=f"{name_prefix}_PortRing", material="vacuum")
            port_hole = self.hfss.modeler.create_circle(
                orientation="XY", origin=[x_feed, y_feed, -Lp_val],
                radius=a_val, name=f"{name_prefix}_PortHole", material="vacuum")
            self.hfss.modeler.subtract(port_ring, [port_hole], keep_originals=False)

            # linha de integração radial inteiramente dentro do anel
            eps_line = min(0.1 * (b_val - a_val), 0.05)
            r_start = a_val + eps_line
            r_end = b_val - eps_line
            if r_end <= r_start:
                r_end = a_val + 0.75 * (b_val - a_val)
            p1 = [x_feed + r_start, y_feed, -Lp_val]
            p2 = [x_feed + r_end,   y_feed, -Lp_val]

            _ = self.hfss.lumped_port(
                assignment=port_ring.name,
                integration_line=[p1, p2],
                impedance=50.0,
                name=f"{name_prefix}_Lumped",
                renormalize=True
            )
            self.log_message(f"Lumped Port '{name_prefix}_Lumped' created (integration line).")
            return pin, None, shield
        except Exception as e:
            self.log_message(f"Exception in coax creation '{name_prefix}': {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            return None, None, None

    # ---------- Inicializa variáveis Pow_/Phi_ ----------
    def _init_postproc_vars(self, excitations: List[str]):
        try:
            n = len(excitations)
            if n <= 0:
                return
            self.update_ports_table(n)
            for i in range(n):
                self.port_params[i]["amp"] = 1.0
                self.port_params[i]["phase"] = 0.0
                idx = i + 1
                self.hfss[f"Pow_P{idx}"] = "1W"
                self.hfss[f"Phi_P{idx}"] = "0deg"
            self.log_message(f"Initialized post-processing vars for {n} ports (Pow_Pi=1W, Phi_Pi=0deg).")
        except Exception as e:
            self.log_message(f"Error initializing post-processing variables: {e}")

    # ---------- Infinite Sphere ----------
    def _create_infinite_sphere(self, name="FF_Sphere"):
        try:
            if hasattr(self.hfss, "create_infinite_sphere"):
                self.hfss.create_infinite_sphere(
                    name=name,
                    theta_start=-180, theta_stop=180, theta_step=1,
                    phi_start=-180,   phi_stop=180,   phi_step=1
                )
                self.log_message(f"Infinite Sphere '{name}' created: θ[-180,180] φ[-180,180] step 1°")
                return name
        except Exception as e:
            self.log_message(f"Infinite Sphere API #1 failed: {e}")
        try:
            if hasattr(self.hfss, "create_infinite_sphere"):
                self.hfss.create_infinite_sphere(
                    name=name,
                    theta_start=0, theta_stop=180, theta_step=1,
                    phi_start=-180, phi_stop=180, phi_step=1
                )
                self.log_message(f"Infinite Sphere '{name}' created with θ[0,180] (fallback).")
                return name
        except Exception as e:
            self.log_message(f"Infinite Sphere API fallback failed: {e}")
        self.log_message("Could not create Infinite Sphere (far-field reports may be unavailable).")
        return None

    # ---------- Far-field cuts ----------
    def _get_gain_cut(self, sphere_name: str, freq_ghz: float, cut: str, fixed_angle_deg: float):
        expr = "GainTotal"
        fstr = f"{freq_ghz}GHz"
        try:
            rpt = self.hfss.post.reports_by_category.radiation(expressions=[expr])
            if cut.lower() == "theta":
                rpt.primary_sweep = "Theta"
                try:
                    rpt.others = {"Phi": [f"{fixed_angle_deg}deg"]}
                except Exception:
                    pass
            else:
                rpt.primary_sweep = "Phi"
                try:
                    rpt.others = {"Theta": [f"{fixed_angle_deg}deg"]}
                except Exception:
                    pass
            try:
                rpt.frequencies = [fstr]
            except Exception:
                pass
            try:
                rpt.context = {"Sphere": sphere_name, "Context": "Infinite Sphere"}
            except Exception:
               sol = rpt.get_solution_data()
            if sol:
                x = np.asarray(sol.primary_sweep_values, dtype=float)
                ydata = sol.data_real()
                y = np.asarray(ydata[0] if isinstance(ydata, (list, tuple)) else ydata, dtype=float)
                return x, y
        except Exception as e:
            self.log_message(f"Radiation report failed for {cut}: {e}")

        try:
            getter = getattr(self.hfss.post, "get_far_field_data", None)
            if getter:
                ff = getter(sphere_name=sphere_name, freq=fstr)
                th = np.asarray(ff.get("theta", []), dtype=float)
                ph = np.asarray(ff.get("phi", []), dtype=float)
                g  = np.asarray(ff.get("gain_total", []), dtype=float)
                if cut.lower() == "theta":
                    mask = np.isclose(ph, fixed_angle_deg, atol=1e-6)
                    if mask.any():
                        return th[mask], g[mask]
                else:
                    mask = np.isclose(th, fixed_angle_deg, atol=1e-6)
                    if mask.any():
                        return ph[mask], g[mask]
        except Exception as e:
            self.log_message(f"get_far_field_data fallback failed: {e}")
        return None, None

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

            # abrir/criar projeto + design
            self._open_or_create_project()
            self.progress_bar.set(0.25)

            self.hfss.modeler.model_units = "mm"
            self.log_message("Model units set to: mm")

            # materiais
            sub_name = self.params["substrate_material"]
            if not self.hfss.materials.checkifmaterialexists(sub_name):
                sub_name = "Custom_Substrate"
                self._ensure_material(sub_name, float(self.params["er"]), float(self.params["tan_d"]))

            # Geometria e variáveis
            L = float(self.calculated_params["patch_length"])
            W = float(self.calculated_params["patch_width"])
            spacing = float(self.calculated_params["spacing"])
            rows = int(self.calculated_params["rows"])
            cols = int(self.calculated_params["cols"])
            h_sub = float(self.params["substrate_thickness"])
            sub_w = float(self.calculated_params["substrate_width"])
            sub_l = float(self.calculated_params["substrate_length"])

            self._set_design_variables(L, W, spacing, rows, cols, h_sub, sub_w, sub_l)

            # Substrato e Ground
            self.log_message("Creating substrate")
            substrate = self.hfss.modeler.create_box(
                origin=["-subW/2", "-subL/2", 0],
                sizes=["subW", "subL", "h_sub"],
                name="Substrate",
                material=sub_name
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
            total_w = cols * W + (cols - 1) * spacing
            total_l = rows * L + (rows - 1) * spacing
            start_x = -total_w / 2 + W / 2
            start_y = -total_l / 2 + L / 2
            self.progress_bar.set(0.35)

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

                    # feed dentro do patch
                    if self.params["feed_position"] == "edge":
                        y_feed = cy - 0.5 * L + 0.02 * L
                    else:
                        y_feed = cy - 0.5 * L + 0.30 * L
                    relx = float(self.params["feed_rel_x"])
                    relx = min(max(relx, 0.0), 1.0)
                    x_feed = cx - 0.5 * W + relx * W

                    pad = self.hfss.modeler.create_circle(
                        orientation="XY",
                        origin=[x_feed, y_feed, "h_sub"],
                        radius="a",
                        name=f"{patch_name}_Pad",
                        material="copper"
                    )
                    try:
                        self.hfss.modeler.unite([patch, pad])
                    except Exception:
                        pass

                    self._create_coax_feed_lumped(
                        ground=ground,
                        substrate=substrate,
                        x_feed=x_feed,
                        y_feed=y_feed,
                        name_prefix=f"P{count}"
                    )
                    self.progress_bar.set(0.35 + 0.25 * (count / float(rows * cols)))

            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return

            try:
                names = [ground.name] + [p.name for p in patches]
                self.hfss.assign_perfecte_to_sheets(names)
                self.log_message(f"PerfectE assigned to: {names}")
            except Exception as e:
                self.log_message(f"PerfectE assignment warning: {e}")

            # região de ar + rad
            self.log_message("Creating air region + radiation boundary")
            lambda0_mm = self.c / (self.params["sweep_start"] * 1e9) * 1000.0
            pad_mm = float(lambda0_mm) / 4.0
            region = self.hfss.modeler.create_region(
                [pad_mm, pad_mm, pad_mm, pad_mm, pad_mm, pad_mm], is_percentage=False
            )
            self.hfss.assign_radiation_boundary_to_objects(region)
            self.progress_bar.set(0.60)

            # Setup + sweep
            self.log_message("Creating simulation setup")
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02

            self.log_message("Creating frequency sweep (Fast)")
            try:
                setup.create_frequency_sweep(
                    unit="GHz",
                    name="Sweep1",
                    start_frequency=self.params["sweep_start"],
                    stop_frequency=self.params["sweep_stop"],
                    sweep_type="Fast"
                )
            except Exception:
                setup.create_frequency_sweep(
                    unit="GHz",
                    name="Sweep1",
                    start_frequency=self.params["sweep_start"],
                    stop_frequency=self.params["sweep_stop"],
                    sweep_type="Interpolating"
                )

            # malha leve
            self.log_message("Assigning local mesh refinement")
            try:
                lambda_g_mm = max(1e-6, self.calculated_params["lambda_g"])
                edge_len = max(lambda_g_mm / 60.0, L / 200.0)
                for p in patches:
                    self.hfss.mesh.assign_length_mesh([p], maximum_length=f"{edge_len}mm")
            except Exception as e:
                self.log_message(f"Mesh refinement warning: {e}")

            # excitações
            try:
                exs = self.hfss.get_excitations_name() or []
            except Exception:
                exs = list(getattr(self.hfss, "excitations", []) or [])
            self.log_message(f"Excitations created: {len(exs)} -> {exs}")
            if not exs:
                self.sim_status_label.configure(text="No excitations defined")
                self.log_message("No excitations found. Aborting before solve.")
                return

            # variáveis de pós-processamento (1W, 0deg)
            self._init_postproc_vars(exs)

            # Infinite Sphere
            sphere_name = self._create_infinite_sphere(name="FF_Sphere")

            self.log_message("Validating design")
            try:
                _ = self.hfss.validate_full_design()
            except Exception as e:
                self.log_message(f"Validation warning: {e}")

            self.log_message("Starting analysis")
            if self.save_project:
                self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])

            if self.stop_simulation:
                self.log_message("Simulation stopped by user")
                return

            self.progress_bar.set(0.9)
            self.log_message("Processing results")
            self._farfield_sphere = sphere_name
            self.plot_results()
            self.progress_bar.set(1.0)
            self.sim_status_label.configure(text="Simulation completed")
            self.log_message("Simulation completed successfully")
            
            # Habilitar botão de ajuste após simulação
            self.adjust_button.configure(state="normal")
        except Exception as e:
            self.log_message(f"Error in simulation: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")
            self.sim_status_label.configure(text=f"Simulation error: {e}")
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    def setup_simulation_tab(self):
        tab = self.tabview.tab("Simulation")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)

        # Título
        title = ctk.CTkLabel(main, text="Simulation Control", 
                            font=ctk.CTkFont(size=18, weight="bold", family="Segoe UI"))
        title.pack(pady=10)

        # Container para botões
        btn_container = ctk.CTkFrame(main, fg_color="transparent")
        btn_container.pack(pady=20)
        
        self.run_button = ctk.CTkButton(btn_container, text="Run Simulation", 
                                       command=self.start_simulation_thread,
                                       fg_color=("#2E8B57", "#2E8B57"), 
                                       hover_color=("#3CB371", "#3CB371"), 
                                       height=45, width=180,
                                       font=ctk.CTkFont(size=14, weight="bold"))
        self.run_button.pack(side="left", padx=15)
        
        self.stop_button = ctk.CTkButton(btn_container, text="Stop Simulation", 
                                        command=self.stop_simulation_thread,
                                        fg_color=("#DC143C", "#DC143C"), 
                                        hover_color=("#FF4500", "#FF4500"),
                                        state="disabled", height=45, width=180,
                                        font=ctk.CTkFont(size=14, weight="bold"))
        self.stop_button.pack(side="left", padx=15)

        # Barra de progresso
        progress_frame = ctk.CTkFrame(main, fg_color="transparent")
        progress_frame.pack(fill="x", padx=50, pady=20)
        
        ctk.CTkLabel(progress_frame, text="Simulation Progress:", 
                    font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(0, 10))
        
        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=20, 
                                              progress_color=("#3B8ED0", "#1F6AA5"))
        self.progress_bar.pack(fill="x", pady=5)
        self.progress_bar.set(0)

        # Status da simulação
        self.sim_status_label = ctk.CTkLabel(main, text="Simulation not started",
                                            font=ctk.CTkFont(weight="bold", size=14))
        self.sim_status_label.pack(pady=15)

        # Nota
        note = ctk.CTkFrame(main, fg_color=("gray90", "gray15"), corner_radius=8)
        note.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(note, text="Note: Simulation time increases with array size.",
                    font=ctk.CTkFont(size=12, slant="italic"),
                    text_color=("gray40", "gray60")).pack(padx=15, pady=15)

    # ---------------- Results Tab / Plots ----------------
    def setup_results_tab(self):
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        # Título
        title = ctk.CTkLabel(main, text="Simulation Results", 
                            font=ctk.CTkFont(size=18, weight="bold", family="Segoe UI"))
        title.grid(row=0, column=0, pady=10)

        # Container dos gráficos
        graph_container = ctk.CTkFrame(main, fg_color=("gray95", "gray20"), corner_radius=8)
        graph_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        graph_container.grid_columnconfigure(0, weight=1)
        graph_container.grid_rowconfigure(0, weight=1)

        # Configuração dos gráficos com tema escuro
        plt.style.use('dark_background')
        self.fig, (self.ax_s11, self.ax_theta, self.ax_phi) = plt.subplots(1, 3, figsize=(14, 5))
        
        # Ajustar cores para modo escuro
        if ctk.get_appearance_mode() == "Dark":
            self.fig.patch.set_facecolor('#2B2B2B')
            for ax in (self.ax_s11, self.ax_theta, self.ax_phi):
                ax.set_facecolor('#2B2B2B')
                ax.tick_params(colors='white')
                ax.xaxis.label.set_color('white')
                ax.yaxis.label.set_color('white')
                ax.title.set_color('white')
                for spine in ax.spines.values():
                    spine.set_color('white')
                ax.grid(color='gray')
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_container)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)

        # Botões de exportação
        export_frame = ctk.CTkFrame(main, fg_color="transparent")
        export_frame.grid(row=2, column=0, pady=15)
        
        ctk.CTkButton(export_frame, text="Export CSV", command=self.export_csv,
                      fg_color=("#6A5ACD", "#6A5ACD"), hover_color=("#7B68EE", "#7B68EE"),
                      height=35, width=120).pack(side="left", padx=10)
        
        ctk.CTkButton(export_frame, text="Export PNG", command=self.export_png,
                      fg_color=("#20B2AA", "#20B2AA"), hover_color=("#40E0D0", "#40E0D0"),
                      height=35, width=120).pack(side="left", padx=10)
        
        # Botão de ajuste de frequência
        self.adjust_button = ctk.CTkButton(export_frame, text="Adjust Frequency", command=self.adjust_frequency,
                                          fg_color=("#FF8C00", "#FF8C00"), hover_color=("#FFA500", "#FFA500"),
                                          height=35, width=120, state="disabled")
        self.adjust_button.pack(side="left", padx=10)

    def plot_results(self):
        try:
            self.log_message("Plotting results")

            # --- S11 ---
            self.ax_s11.clear()
            try:
                exs = self.hfss.get_excitations_name() or []
            except Exception:
                exs = []
            expr = "dB(S(1,1))"
            if exs:
                p = exs[0].split(":")[0]
                expr = f"dB(S({p},{p}))"
            rpt = self.hfss.post.reports_by_category.standard(expressions=[expr])
            rpt.context = ["Setup1: Sweep1"]
            sol = rpt.get_solution_data()

            if sol:
                freqs = np.asarray(sol.primary_sweep_values, dtype=float)
                data = sol.data_real()
                if isinstance(data, (list, tuple)) and len(data) > 0 and hasattr(data[0], "__len__"):
                    y = np.asarray(data[0], dtype=float)
                else:
                    y = np.asarray(data, dtype=float)
                if y.size == freqs.size:
                    self.simulation_data = np.column_stack((freqs, y))
                    self.ax_s11.plot(freqs, y, linewidth=2, label=expr)
                    self.ax_s11.axhline(y=-10, linestyle='--', alpha=0.7, label='-10 dB')
                    self.ax_s11.set_xlabel("Frequency (GHz)")
                    self.ax_s11.set_ylabel("S-Parameter (dB)")
                    self.ax_s11.set_title("S11")
                    self.ax_s11.legend()
                    self.ax_s11.grid(True)
                    cf = float(self.params["frequency"])
                    self.ax_s11.axvline(x=cf, linestyle='--', alpha=0.7)
                    
                    # Encontrar a frequência de ressonância (mínimo de S11)
                    min_idx = np.argmin(y)
                    self.simulated_resonance_freq = freqs[min_idx]
                    self.ax_s11.axvline(x=self.simulated_resonance_freq, linestyle='--', color='red', alpha=0.7)
                    self.ax_s11.text(self.simulated_resonance_freq, np.min(y)-2, 
                                    f'{self.simulated_resonance_freq:.2f} GHz', 
                                    color='red', ha='center')
            else:
                self.log_message("Could not get S11 data")

            # --- Far-field cuts @ f0 ---
            self.ax_theta.clear()
            self.ax_phi.clear()
            f0 = float(self.params["frequency"])
            sphere = getattr(self, "_farfield_sphere", "FF_Sphere")

            # Gain vs Theta @ phi=0°
            th, gth = self._get_gain_cut(sphere, f0, cut="theta", fixed_angle_deg=0.0)
            if th is not None and gth is not None:
                self.ax_theta.plot(th, gth, linewidth=2)
                self.ax_theta.set_xlabel("Theta (deg)")
                self.ax_theta.set_ylabel("GainTotal (dB)")
                self.ax_theta.set_title("Gain vs Theta @ Phi=0°")
                self.ax_theta.grid(True)
            else:
                self.log_message("Theta-cut gain not available (check Infinite Sphere / version).")

            # Gain vs Phi @ theta=90°
            ph, gph = self._get_gain_cut(sphere, f0, cut="phi", fixed_angle_deg=90.0)
            if ph is not None and gph is not None:
                self.ax_phi.plot(ph, gph, linewidth=2)
                self.ax_phi.set_xlabel("Phi (deg)")
                self.ax_phi.set_ylabel("GainTotal (dB)")
                self.ax_phi.set_title("Gain vs Phi @ Theta=90°")
                self.ax_phi.grid(True)
            else:
                self.log_message("Phi-cut gain not available (check Infinite Sphere / version).")

            self.canvas.draw()
            self.log_message("Results plotted successfully")
        except Exception as e:
            self.log_message(f"Error plotting results: {e}")
            self.log_message(f"Traceback: {traceback.format_exc()}")

    def adjust_frequency(self):
        """Ajusta a frequência com base na diferença entre a ressonância simulada e a especificada."""
        if not hasattr(self, 'simulated_resonance_freq') or self.simulated_resonance_freq is None:
            self.log_message("No simulation data available for frequency adjustment")
            return
            
        specified_freq = self.params["frequency"]
        simulated_freq = self.simulated_resonance_freq
        delta_f = (specified_freq - simulated_freq) * 1000  # Diferença em MHz
        
        # Perguntar ao usuário se deseja ajustar
        response = messagebox.askyesno(
            "Frequency Adjustment",
            f"Specified frequency: {specified_freq:.3f} GHz\n"
            f"Simulated resonance: {simulated_freq:.3f} GHz\n"
            f"Difference: {delta_f:.2f} MHz\n\n"
            "Do you want to adjust the parameters to correct the frequency?"
        )
        
        if response:
            # Calcular fator de ajuste baseado na relação de comprimento de onda
            # λ_simulado / λ_especificado = f_especificado / f_simulado
            # Queremos ajustar as dimensões por este fator
            scaling_factor = specified_freq / simulated_freq
            
            # Ajustar a frequência para o valor especificado
            # e recalcular todos os parâmetros dependentes de frequência
            self.log_message(f"Adjusting parameters with scaling factor: {scaling_factor:.4f}")
            
            # Atualizar a frequência na interface
            for key, widget in self.entries:
                if key == "frequency":
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end")
                        widget.insert(0, str(specified_freq))
            
            # Recalcular todos os parâmetros
            self.calculate_parameters()
            
            # Executar nova simulação
            self.start_simulation_thread()

    def setup_log_tab(self):
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        # Título
        title = ctk.CTkLabel(main, text="Simulation Log", 
                            font=ctk.CTkFont(size=18, weight="bold", family="Segoe UI"))
        title.grid(row=0, column=0, pady=10)

        # Área de texto do log
        self.log_text = ctk.CTkTextbox(main, width=900, height=500, 
                                      font=ctk.CTkFont(family="Consolas", size=12),
                                      fg_color=("gray95", "gray20"),
                                      border_color=("#3B8ED0", "#1F6AA5"),
                                      border_width=1)
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")

        # Botões de ação do log
        btn_frame = ctk.CTkFrame(main, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=10)
        
        ctk.CTkButton(btn_frame, text="Clear Log", command=self.clear_log,
                      fg_color=("gray70", "gray40"), hover_color=("gray60", "gray50"),
                      height=32, width=100).pack(side="left", padx=10)
        
        ctk.CTkButton(btn_frame, text="Save Log", command=self.save_log,
                      fg_color=("#3B8ED0", "#1F6AA5"), hover_color=("#1F6AA5", "#1F538D"),
                      height=32, width=100).pack(side="left", padx=10)

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
            if hasattr(self, 'simulation_data'):
                np.savetxt("simulation_results.csv", self.simulation_data, delimiter=",",
                           header="Frequency (GHz), S11 (dB)", comments='')
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

    # Encerramento / IO
    def cleanup(self):
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
        self.log_message("Application closing...")
        self.cleanup()
        self.window.quit()
        self.window.destroy()

    def save_parameters(self):
        try:
            self._sync_port_params_from_ui()
            all_params = {**self.params, **self.calculated_params, "port_params": self.port_params}
            with open("antenna_parameters.json", "w") as f:
                json.dump(all_params, f, indent=4)
            self.log_message("Parameters saved to antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error saving parameters: {e}")

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
            if "port_params" in all_params and isinstance(all_params["port_params"], list):
                self.port_params = all_params["port_params"]
            self.update_interface_from_params()
            self.log_message("Parameters loaded from antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error loading parameters: {e}")

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
            self.rows_cols_label.configure(text=f"Configuration: {self.calculated_params['rows']} x "
                                                f"{self.calculated_params['cols']}")
            self.spacing_label.configure(text=f"Spacing: {self.calculated_params['spacing']:.2f} mm "
                                              f"({self.params['spacing_type']})")
            self.dimensions_label.configure(
                text=f"Patch Dimensions: {self.calculated_params['patch_length']:.2f} x "
                     f"{self.calculated_params['patch_width']:.2f} mm")
            self.lambda_label.configure(text=f"Guided Wavelength: {self.calculated_params['lambda_g']:.2f} mm")
            self.feed_offset_label.configure(text=f"Feed Offset (y): {self.calculated_params['feed_offset']:.2f} mm")
            self.substrate_dims_label.configure(
                text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x "
                     f"{self.calculated_params['substrate_length']:.2f} mm")

            self.update_ports_table(self.calculated_params["num_patches"])
            for i, (e_amp, e_ph) in enumerate(self.port_entries):
                e_amp.delete(0, "end"); e_ph.delete(0, "end")
                e_amp.insert(0, str(self.port_params[i]["amp"]))
                e_ph.insert(0, str(self.port_params[i]["phase"]))
            self.log_message("Interface updated with loaded parameters")
        except Exception as e:
            self.log_message(f"Error updating interface: {e}")

    def run(self):
        try:
            self.window.mainloop()
        finally:
            self.cleanup()


if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()