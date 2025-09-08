# -*- coding: utf-8 -*-
"""
Modern Patch Antenna Designer (v3.1 - Slot-Coupled with Fix)
------------------------------------------------------------
Aprimoramentos chave em relação à v3:
- Corrigido TypeError na função create_section ao passar 'columnspan'.
- A função agora aceita **kwargs para maior flexibilidade no layout.

Observação: Este script assume um ambiente com Ansys Electronics Desktop (AEDT)
instalado/licenciado e PyAEDT compatível disponível.
"""

import os
import re
import tempfile
from datetime import datetime
import math
import json
import traceback
import queue
import threading
from typing import Tuple, List, Optional, Dict

import numpy as np
import matplotlib
matplotlib.use("TkAgg")  # Necessário para embutir em Tk
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401 (necessário para mplot3d)
import customtkinter as ctk

from ansys.aedt.core import Desktop, Hfss

# ---------- Aparência ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ModernPatchAntennaDesigner:
    """Aplicativo GUI para dimensionamento e simulação de array de antenas com acoplamento por fenda em HFSS.

    Fluxo principal:
        1) Usuário define parâmetros e clica em "Calculate Parameters".
        2) Usuário executa a simulação (tab Simulation) -> geometria + setup + solve.
        3) Pós-processamento aplica fontes iniciais (1W/0°) e disponibiliza UI para beamforming.
        4) Resultados: S11/VSWR/|Z|, cortes de ganho e superfície 3D.
    """

    # ---------------- Inicialização ----------------
    def __init__(self):
        # AEDT
        self.hfss: Optional[Hfss] = None
        self.desktop: Optional[Desktop] = None
        self.temp_folder = None
        self.project_path = ""
        self.project_display_name = "slot_coupled_array"
        self.design_base_name = "slot_coupled_array"

        # Runtime
        self.log_queue = queue.Queue()
        self.is_simulation_running = False
        self.save_project = False
        self.stop_simulation = False
        self.created_ports: List[str] = []

        # Dados em memória
        self.last_s11_analysis = None  # dict: f, s11_db, |Z|, f_res, Z(f_res) etc
        self.theta_cut = None        # (theta, gain)
        self.phi_cut = None          # (phi, gain)
        self.grid3d = None           # (TH, PH, Gdb)
        self.auto_refresh_job = None

        # Parâmetros do usuário (default)
        self.params = {
            "frequency": 10.0,
            "gain": 12.0,
            "sweep_start": 8.0,
            "sweep_stop": 12.0,
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,
            "spacing_type": "0.8*lambda",
            # Patch Substrate
            "patch_substrate_material": "Duroid (tm)",
            "patch_substrate_thickness": 0.5,
            "er_patch": 2.2,
            "tan_d_patch": 0.0009,
            "metal_thickness": 0.035,
            # Feed Substrate
            "feed_substrate_material": "Rogers RO4003C (tm)",
            "feed_substrate_thickness": 0.5,
            "er_feed": 3.55,
            "tan_d_feed": 0.0027,
            # Stack-up
            "air_gap_thickness": 5.0,
            "honeycomb_thickness": 5.0,
            "honeycomb_material": "Air",
            # Slot & Dipole
            "slot_length": 8.0,
            "slot_width": 0.5,
            "slot_separation": 4.0, # Distância entre centros das fendas
            "dipole_length": 9.0,
            "dipole_width": 0.5,
            "dipole_gap": 0.5,
            # Sim settings
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
        """Constroi a janela principal e abas."""
        self.window = ctk.CTk()
        self.window.title("Slot-Coupled Patch Antenna Array Designer")
        try:
            self.window.state("zoomed")
        except Exception:
            self.window.geometry("1500x950")

        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.window, height=70, fg_color=("gray85", "gray20"))
        header.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        header.grid_propagate(False)
        ctk.CTkLabel(
            header, text="Slot-Coupled Patch Antenna Array Designer",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=("gray10", "gray90")
        ).pack(pady=18)

        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        for name in ["Design Parameters", "Simulation", "Results", "Log"]:
            self.tabview.add(name)
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()

        status = ctk.CTkFrame(self.window, height=38)
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        self.status_label = ctk.CTkLabel(status, text="Ready", font=ctk.CTkFont(weight="bold"))
        self.status_label.pack(pady=6)

        self.process_log_queue()

    def create_section(self, parent, title, row, column, padx=10, pady=10, **kwargs):
        """
        Cria um *frame* com título e separador para organizar a UI.
        Aceita **kwargs para passar ao método .grid() (ex: columnspan).
        """
        section = ctk.CTkFrame(parent, fg_color=("gray92", "gray18"))
        section.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady, **kwargs)
        section.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(section, text=title, font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=("gray20", "gray80")).grid(row=0, column=0, sticky="w", padx=15, pady=(10, 6))
        ctk.CTkFrame(section, height=2, fg_color=("gray70", "gray30")).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        return section

    def setup_parameters_tab(self):
        """Aba de parâmetros de projeto/substrato/alimentação/simulação."""
        tab = self.tabview.tab("Design Parameters")
        tab.grid_columnconfigure((0, 1), weight=1)
        tab.grid_rowconfigure(0, weight=1)
        main_frame = ctk.CTkScrollableFrame(tab)
        main_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=6, pady=6)
        main_frame.grid_columnconfigure((0, 1), weight=1)

        self.entries = []
        
        def add_entry(section, label, key, value, row, col, combo=None, check=False):
            frame = ctk.CTkFrame(section)
            frame.grid(row=row, column=col, padx=5, pady=2, sticky="ew")
            ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(weight="bold")
                         ).pack(side="left", padx=(10, 5))
            
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(frame, values=combo, variable=var, width=200)
                widget.pack(side="right", padx=(5, 10))
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(frame, text="", variable=var)
                widget.pack(side="right", padx=(5, 10))
                self.entries.append((key, var))
            else:
                widget = ctk.CTkEntry(frame, width=200)
                widget.insert(0, str(value))
                widget.pack(side="right", padx=(5, 10))
                self.entries.append((key, widget))
            return row + (col % 2)

        # Coluna 1
        sec_ant = self.create_section(main_frame, "Antenna Parameters", 0, 0)
        sec_ant.grid_columnconfigure(0, weight=1)
        r = 2
        r = add_entry(sec_ant, "Central Frequency (GHz):", "frequency", self.params["frequency"], r, 0)
        r = add_entry(sec_ant, "Desired Gain (dBi):", "gain", self.params["gain"], r, 0)
        r = add_entry(sec_ant, "Sweep Start (GHz):", "sweep_start", self.params["sweep_start"], r, 0)
        r = add_entry(sec_ant, "Sweep Stop (GHz):", "sweep_stop", self.params["sweep_stop"], r, 0)
        r = add_entry(sec_ant, "Patch Spacing:", "spacing_type", self.params["spacing_type"], r, 0,
                      combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        sec_patch_sub = self.create_section(main_frame, "Patch Substrate", 1, 0)
        sec_patch_sub.grid_columnconfigure(0, weight=1)
        r = 2
        r = add_entry(sec_patch_sub, "Material:", "patch_substrate_material", self.params["patch_substrate_material"], r, 0,
                      combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        r = add_entry(sec_patch_sub, "Rel. Permittivity (εr):", "er_patch", self.params["er_patch"], r, 0)
        r = add_entry(sec_patch_sub, "Loss Tangent (tan δ):", "tan_d_patch", self.params["tan_d_patch"], r, 0)
        r = add_entry(sec_patch_sub, "Thickness (mm):", "patch_substrate_thickness", self.params["patch_substrate_thickness"], r, 0)
        r = add_entry(sec_patch_sub, "Metal Thickness (mm):", "metal_thickness", self.params["metal_thickness"], r, 0)

        sec_sim = self.create_section(main_frame, "Simulation Settings", 2, 0)
        sec_sim.grid_columnconfigure(0, weight=1)
        r = 2
        r = add_entry(sec_sim, "CPU Cores:", "cores", self.params["cores"], r, 0)
        r = add_entry(sec_sim, "Show HFSS Interface:", "show_gui", not self.params["non_graphical"], r, 0, check=True)
        r = add_entry(sec_sim, "Save Project:", "save_project", self.save_project, r, 0, check=True)
        r = add_entry(sec_sim, "Sweep Type:", "sweep_type", self.params["sweep_type"], r, 0,
                      combo=["Discrete", "Interpolating", "Fast"])
        r = add_entry(sec_sim, "Discrete Step (GHz):", "sweep_step", self.params["sweep_step"], r, 0)
        r = add_entry(sec_sim, "3D Theta step (deg):", "theta_step", self.params["theta_step"], r, 0)
        r = add_entry(sec_sim, "3D Phi step (deg):", "phi_step", self.params["phi_step"], r, 0)

        # Coluna 2
        sec_feed_sub = self.create_section(main_frame, "Feed Substrate", 0, 1)
        sec_feed_sub.grid_columnconfigure(0, weight=1)
        r=2
        r = add_entry(sec_feed_sub, "Material:", "feed_substrate_material", self.params["feed_substrate_material"], r, 0,
                      combo=["Rogers RO4003C (tm)", "Duroid (tm)", "FR4_epoxy", "Air"])
        r = add_entry(sec_feed_sub, "Rel. Permittivity (εr):", "er_feed", self.params["er_feed"], r, 0)
        r = add_entry(sec_feed_sub, "Loss Tangent (tan δ):", "tan_d_feed", self.params["tan_d_feed"], r, 0)
        r = add_entry(sec_feed_sub, "Thickness (mm):", "feed_substrate_thickness", self.params["feed_substrate_thickness"], r, 0)

        sec_stackup = self.create_section(main_frame, "Stack-up Layers", 1, 1)
        sec_stackup.grid_columnconfigure(0, weight=1)
        r=2
        r = add_entry(sec_stackup, "Air Gap Thickness (mm):", "air_gap_thickness", self.params["air_gap_thickness"], r, 0)
        r = add_entry(sec_stackup, "Honeycomb Material:", "honeycomb_material", self.params["honeycomb_material"], r, 0, combo=["Air", "Nomex"])
        r = add_entry(sec_stackup, "Honeycomb Thickness (mm):", "honeycomb_thickness", self.params["honeycomb_thickness"], r, 0)

        sec_feed = self.create_section(main_frame, "Slot & Dipole Feed", 2, 1)
        sec_feed.grid_columnconfigure(0, weight=1)
        r=2
        r = add_entry(sec_feed, "Slot Length (mm):", "slot_length", self.params["slot_length"], r, 0)
        r = add_entry(sec_feed, "Slot Width (mm):", "slot_width", self.params["slot_width"], r, 0)
        r = add_entry(sec_feed, "Slot Separation (mm):", "slot_separation", self.params["slot_separation"], r, 0)
        r = add_entry(sec_feed, "Dipole Length (mm):", "dipole_length", self.params["dipole_length"], r, 0)
        r = add_entry(sec_feed, "Dipole Width (mm):", "dipole_width", self.params["dipole_width"], r, 0)
        r = add_entry(sec_feed, "Dipole Gap (mm):", "dipole_gap", self.params["dipole_gap"], r, 0)
        
        # Calculated Parameters (abaixo das colunas)
        sec_calc = self.create_section(main_frame, "Calculated Parameters", 3, 0, columnspan=2)
        grid = ctk.CTkFrame(sec_calc); grid.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        grid.columnconfigure((0, 1), weight=1)
        self.patches_label = ctk.CTkLabel(grid, text="Number of Patches: 4", font=ctk.CTkFont(weight="bold"))
        self.patches_label.grid(row=0, column=0, sticky="w", pady=4)
        self.rows_cols_label = ctk.CTkLabel(grid, text="Configuration: 2 x 2", font=ctk.CTkFont(weight="bold"))
        self.rows_cols_label.grid(row=0, column=1, sticky="w", pady=4)
        self.spacing_label = ctk.CTkLabel(grid, text="Spacing: -- mm", font=ctk.CTkFont(weight="bold"))
        self.spacing_label.grid(row=1, column=0, sticky="w", pady=4)
        self.dimensions_label = ctk.CTkLabel(grid, text="Patch Dimensions: -- x -- mm", font=ctk.CTkFont(weight="bold"))
        self.dimensions_label.grid(row=1, column=1, sticky="w", pady=4)
        self.lambda_label = ctk.CTkLabel(grid, text="Guided Wavelength: -- mm", font=ctk.CTkFont(weight="bold"))
        self.lambda_label.grid(row=2, column=0, sticky="w", pady=4)
        self.feed_offset_label = ctk.CTkLabel(grid, text="Feed Offset (y): -- mm", font=ctk.CTkFont(weight="bold"))
        self.feed_offset_label.grid(row=2, column=1, sticky="w", pady=4)
        self.substrate_dims_label = ctk.CTkLabel(grid, text="Substrate Dimensions: -- x -- mm",
                                                 font=ctk.CTkFont(weight="bold"))
        self.substrate_dims_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=4)

        btns = ctk.CTkFrame(sec_calc); btns.grid(row=3, column=0, sticky="ew", padx=15, pady=12)
        ctk.CTkButton(btns, text="Calculate Parameters", command=self.calculate_parameters,
                      fg_color="#2E8B57", hover_color="#3CB371", width=180).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="Save Parameters", command=self.save_parameters,
                      fg_color="#4169E1", hover_color="#6495ED", width=140).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="Load Parameters", command=self.load_parameters,
                      fg_color="#FF8C00", hover_color="#FFA500", width=140).pack(side="left", padx=8)

    def setup_simulation_tab(self):
        """Aba com botões de execução/parada e barra de progresso."""
        tab = self.tabview.tab("Simulation")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(main, text="Simulation Control", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        row = ctk.CTkFrame(main); row.pack(pady=14)
        self.run_button = ctk.CTkButton(row, text="Run Simulation", command=self.start_simulation_thread,
                                         fg_color="#2E8B57", hover_color="#3CB371", height=40, width=160)
        self.run_button.pack(side="left", padx=8)
        self.stop_button = ctk.CTkButton(row, text="Stop Simulation", command=self.stop_simulation_thread,
                                         fg_color="#DC143C", hover_color="#FF4500",
                                         state="disabled", height=40, width=160)
        self.stop_button.pack(side="left", padx=8)
        bar = ctk.CTkFrame(main); bar.pack(fill="x", padx=50, pady=8)
        ctk.CTkLabel(bar, text="Simulation Progress:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.progress_bar = ctk.CTkProgressBar(bar, height=18); self.progress_bar.pack(fill="x", pady=6); self.progress_bar.set(0)
        self.sim_status_label = ctk.CTkLabel(main, text="Simulation not started", font=ctk.CTkFont(weight="bold"))
        self.sim_status_label.pack(pady=8)
        note = ctk.CTkFrame(main, fg_color=("gray90", "gray15")); note.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(note, text="Tip: With direct sources (Magnitude/Phase) you can retune beams without re-solving.",
                     font=ctk.CTkFont(size=12, slant="italic"), text_color=("gray40", "gray60")).pack(padx=10, pady=10)

    def setup_results_tab(self):
        """Aba de resultados com 5 painéis de gráficos + controles."""
        tab = self.tabview.tab("Results")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(main, text="Results & Beamforming", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=10)

        # Área de gráficos com GridSpec 3x2
        graph_frame = ctk.CTkFrame(main)
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
                ax.xaxis.label.set_color('white'); ax.yaxis.label.set_color('white'); ax.title.set_color('white')
                for s in ['bottom', 'top', 'right', 'left']: ax.spines[s].set_color('white')
                ax.grid(color='gray', alpha=0.5)
        self.ax_3d.set_facecolor(face)

        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # Painel de beamforming / refresh
        panel = ctk.CTkFrame(main)
        panel.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        panel.grid_columnconfigure(0, weight=1)
        self.src_frame = ctk.CTkFrame(panel); self.src_frame.grid(row=0, column=0, sticky="ew", padx=6, pady=(8, 2))
        self.source_controls: Dict[str, Dict[str, ctk.CTkBaseClass]] = {}
        self.auto_refresh_var = ctk.BooleanVar(value=False)

        ctrl = ctk.CTkFrame(panel); ctrl.grid(row=1, column=0, pady=6)
        ctk.CTkButton(ctrl, text="Analyze S11", command=self.analyze_and_mark_s11,
                      fg_color="#6A5ACD", hover_color="#7B68EE").pack(side="left", padx=8)
        ctk.CTkButton(ctrl, text="Apply Sources", command=self.apply_sources_from_ui,
                      fg_color="#20B2AA", hover_color="#40E0D0").pack(side="left", padx=8)
        ctk.CTkButton(ctrl, text="Refresh Patterns", command=self.refresh_patterns_only,
                      fg_color="#FF8C00", hover_color="#FFA500").pack(side="left", padx=8)
        ctk.CTkCheckBox(ctrl, text="Auto-refresh (1.5s)", variable=self.auto_refresh_var,
                        command=self.toggle_auto_refresh).pack(side="left", padx=8)
        ctk.CTkButton(ctrl, text="Export PNG", command=self.export_png,
                      fg_color="#20B2AA", hover_color="#40E0D0").pack(side="left", padx=8)
        ctk.CTkButton(ctrl, text="Export CSV (S11)", command=self.export_csv,
                      fg_color="#6A5ACD", hover_color="#7B68EE").pack(side="left", padx=8)

        self.result_label = ctk.CTkLabel(main, text="", font=ctk.CTkFont(weight="bold"))
        self.result_label.grid(row=3, column=0, pady=(6, 2))

    def setup_log_tab(self):
        """Aba de log com *textbox* e botões de limpar/salvar."""
        tab = self.tabview.tab("Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(main, text="Simulation Log", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=10)
        self.log_text = ctk.CTkTextbox(main, width=900, height=500, font=ctk.CTkFont(family="Consolas"))
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", "Log started at " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        btn = ctk.CTkFrame(main); btn.grid(row=2, column=0, pady=8)
        ctk.CTkButton(btn, text="Clear Log", command=self.clear_log).pack(side="left", padx=8)
        ctk.CTkButton(btn, text="Save Log", command=self.save_log).pack(side="left", padx=8)

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
        """Valida faixas de parâmetros mais críticas. Atualiza status/log e retorna True/False."""
        ok = True
        msgs = []
        if self.params["frequency"] <= 0:
            ok = False; msgs.append("frequency must be > 0")
        if self.params["sweep_start"] <= 0 or self.params["sweep_stop"] <= 0:
            ok = False; msgs.append("sweep_start/stop must be > 0")
        if self.params["sweep_start"] >= self.params["sweep_stop"]:
            ok = False; msgs.append("sweep_start must be < sweep_stop")
        if self.params["er_patch"] < 1 or self.params["er_feed"] < 1:
            ok = False; msgs.append("er must be >= 1 for all substrates")
        if self.params["patch_substrate_thickness"] <= 0 or self.params["feed_substrate_thickness"] <= 0:
            ok = False; msgs.append("substrate thickness must be > 0")
        if self.params["slot_length"] <= 0 or self.params["dipole_length"] <= 0:
            ok = False; msgs.append("slot/dipole length must be > 0")
        if self.params["dipole_gap"] >= self.params["dipole_length"]:
            ok = False; msgs.append("dipole gap must be smaller than dipole length")
        if not ok:
            msg = "; ".join(msgs)
            self.status_label.configure(text=f"Invalid parameters: {msg}")
            self.log_message(f"Invalid parameters: {msg}")
        return ok

    def get_parameters(self) -> bool:
        """Lê valores da UI, faz *casting* e sincroniza `self.params`.
        Retorna False se algum valor for inválido.
        """
        self.log_message("Getting parameters from interface")
        for key, widget in self.entries:
            try:
                if key == "cores":
                    self.params[key] = int(widget.get())
                elif key == "show_gui":
                    self.params["non_graphical"] = not widget.get()
                elif key == "save_project":
                    self.save_project = widget.get()
                elif isinstance(widget, (ctk.CTkEntry, ctk.StringVar)):
                     # Check if it should be float or string
                    val = widget.get()
                    if key in ["aedt_version", "spacing_type", "patch_substrate_material", "feed_substrate_material", "honeycomb_material", "sweep_type"]:
                        self.params[key] = str(val)
                    else:
                        self.params[key] = float(val)
            except Exception as e:
                msg = f"Invalid value for {key}: {e}"
                self.status_label.configure(text=msg)
                self.log_message(msg)
                return False
        self.log_message("All parameters retrieved successfully")
        return self._validate_ranges()

    def calculate_patch_dimensions(self, frequency_ghz: float) -> Tuple[float, float, float]:
        """Calcula L, W e λg (em mm) para microfita retangular (usado como estimativa inicial)."""
        f = frequency_ghz * 1e9
        er = float(self.params["er_patch"])
        h = float(self.params["patch_substrate_thickness"]) / 1000.0  # mm->m
        W = self.c / (2 * f) * math.sqrt(2 / (er + 1))
        eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W) ** -0.5
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        L_eff = self.c / (2 * f * math.sqrt(eeff))
        L = L_eff - 2 * dL
        lambda_g = self.c / (f * math.sqrt(eeff))
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)

    def _size_array_from_gain(self) -> Tuple[int, int, int]:
        """Deriva nº de elementos (linhas/colunas) a partir do *gain* desejado."""
        G_elem = 6.0  # Assumed gain for a single slot-coupled patch
        G_des = float(self.params["gain"])
        N_req = max(1, int(math.ceil(10 ** ((G_des - G_elem) / 10.0))))
        if N_req == 1:
            return 1, 1, 1
        
        rows = max(1, int(round(math.sqrt(N_req))))
        cols = max(1, int(math.ceil(N_req / rows)))
        
        return rows, cols, rows * cols

    def calculate_substrate_size(self):
        """Define dimensões do substrato com *margin*."""
        L = self.calculated_params["patch_length"]
        W = self.calculated_params["patch_width"]
        s = self.calculated_params["spacing"]
        r = self.calculated_params["rows"]
        c = self.calculated_params["cols"]
        total_w = c * W + (c - 1) * s
        total_l = r * L + (r - 1) * s
        margin = max(total_w, total_l) * 0.20 + self.params["frequency"] # Margin larger for lower freqs
        self.calculated_params["substrate_width"] = total_w + 2 * margin
        self.calculated_params["substrate_length"] = total_l + 2 * margin
        self.log_message(f"Substrate size calculated: {self.calculated_params['substrate_width']:.2f} x "
                         f"{self.calculated_params['substrate_length']:.2f} mm")

    def calculate_parameters(self):
        """Calcula L/W/λg, *spacing* e *layout* (linhas/colunas). Atualiza UI."""
        self.log_message("Starting parameter calculation")
        if not self.get_parameters():
            self.log_message("Parameter calculation failed due to invalid input")
            return
        try:
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params.update({"patch_length": L_mm, "patch_width": W_mm, "lambda_g": lambda_g_mm})
            lambda0_m = self.c / (self.params["frequency"] * 1e9)
            factors = {"lambda/2": 0.5, "lambda": 1.0, "0.7*lambda": 0.7, "0.8*lambda": 0.8, "0.9*lambda": 0.9}
            spacing_mm = factors.get(self.params["spacing_type"], 0.8) * lambda0_m * 1000.0
            self.calculated_params["spacing"] = spacing_mm
            rows, cols, N_total = self._size_array_from_gain()
            self.calculated_params.update({"num_patches": N_total, "rows": rows, "cols": cols})
            self.log_message(f"Array sizing -> target gain {self.params['gain']} dBi, layout {rows}x{cols} (= {N_total} patches)")
            self.calculate_substrate_size()
            # UI
            self.patches_label.configure(text=f"Number of Patches: {rows*cols}")
            self.rows_cols_label.configure(text=f"Configuration: {rows} x {cols}")
            self.spacing_label.configure(text=f"Spacing: {spacing_mm:.2f} mm ({self.params['spacing_type']})")
            self.dimensions_label.configure(text=f"Patch Dimensions: {L_mm:.2f} x {W_mm:.2f} mm (Initial Estimate)")
            self.lambda_label.configure(text=f"Guided Wavelength: {lambda_g_mm:.2f} mm")
            self.substrate_dims_label.configure(
                text=f"Substrate Dimensions: {self.calculated_params['substrate_width']:.2f} x "
                     f"{self.calculated_params['substrate_length']:.2f} mm")
            self.status_label.configure(text="Parameters calculated successfully")
            self.log_message("Parameters calculated successfully")
        except Exception as e:
            self.status_label.configure(text=f"Error in calculation: {e}")
            self.log_message(f"Error in calculation: {e}\nTraceback: {traceback.format_exc()}")

    # --------- AEDT helpers ---------
    def _ensure_material(self, name: str, er: float, tan_d: float):
        """Garante a existência de um material com εr e tanδ informados."""
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
        self.hfss = Hfss(project=self.project_path, design=self.design_base_name, solution_type="DrivenModal")
        self.log_message(f"Created new project: {self.project_path} (design '{self.design_base_name}')")

    def _set_design_variables(self):
        """Cria/atualiza variáveis do *design* em HFSS (unidades mm/GHz)."""
        self.hfss["f0"] = f"{self.params['frequency']}GHz"
        # Patch substrate
        self.hfss["h_sub"] = f"{self.params['patch_substrate_thickness']}mm"
        self.hfss["t_met"] = f"{self.params['metal_thickness']}mm"
        # Feed substrate
        self.hfss["h_feed_sub"] = f"{self.params['feed_substrate_thickness']}mm"
        # Stackup
        self.hfss["h_air"] = f"{self.params['air_gap_thickness']}mm"
        self.hfss["h_honey"] = f"{self.params['honeycomb_thickness']}mm"
        # Array
        self.hfss["patchL"] = f"{self.calculated_params['patch_length']}mm"
        self.hfss["patchW"] = f"{self.calculated_params['patch_width']}mm"
        self.hfss["spacing"] = f"{self.calculated_params['spacing']}mm"
        self.hfss["rows"] = str(self.calculated_params['rows'])
        self.hfss["cols"] = str(self.calculated_params['cols'])
        self.hfss["subW"] = f"{self.calculated_params['substrate_width']}mm"
        self.hfss["subL"] = f"{self.calculated_params['substrate_length']}mm"
        # Feed
        self.hfss["slotL"] = f"{self.params['slot_length']}mm"
        self.hfss["slotW"] = f"{self.params['slot_width']}mm"
        self.hfss["slotSep"] = f"{self.params['slot_separation']}mm"
        self.hfss["dipoleL"] = f"{self.params['dipole_length']}mm"
        self.hfss["dipoleW"] = f"{self.params['dipole_width']}mm"
        self.hfss["dipoleGap"] = f"{self.params['dipole_gap']}mm"
        
        self.log_message(f"HFSS design variables set.")

    def _create_slotted_coupled_element(self, cx: float, cy: float, name_prefix: str) -> Tuple[list, list, object]:
        """
        Constrói um elemento completo: dipolo de alimentação, fendas de acoplamento e patch radiante.
        Retorna os objetos criados para operações posteriores (ex: subtração).
        """
        # --- Z Coordinates (calculadas a partir das variáveis HFSS) ---
        z_dipole = self.hfss.modeler.get_variable_value("h_feed_sub")
        z_gnd = z_dipole + self.hfss.modeler.get_variable_value("h_air")
        z_patch = z_gnd + self.hfss.modeler.get_variable_value("h_sub")

        # --- Dipole Feed (orientado ao longo do eixo X) ---
        dipole_total_l = self.hfss.modeler.get_variable_value("dipoleL")
        dipole_w = self.hfss.modeler.get_variable_value("dipoleW")
        dipole_gap = self.hfss.modeler.get_variable_value("dipoleGap")
        dipole_arm_l = (dipole_total_l - dipole_gap) / 2.0

        arm1 = self.hfss.modeler.create_rectangle(
            "XY", [cx - dipole_total_l / 2.0, cy - dipole_w / 2.0, z_dipole],
            [dipole_arm_l, dipole_w], name=f"{name_prefix}_Arm1", matname="copper"
        )
        arm2 = self.hfss.modeler.create_rectangle(
            "XY", [cx + dipole_gap / 2.0, cy - dipole_w / 2.0, z_dipole],
            [dipole_arm_l, dipole_w], name=f"{name_prefix}_Arm2", matname="copper"
        )

        # --- Lumped Port ---
        port_sheet = self.hfss.modeler.create_rectangle(
            "YZ", [cx - dipole_gap / 2.0, cy - dipole_w / 2.0, z_dipole],
            [dipole_w, dipole_gap], name=f"{name_prefix}_PortSheet"
        )
        port_name = f"{name_prefix}_Lumped"
        self.hfss.lumped_port(
            assignment=port_sheet.name, impedance=50.0, name=port_name, renormalize=True
        )
        if port_name not in self.created_ports:
            self.created_ports.append(port_name)

        # --- Coupling Slots (ferramentas para subtração posterior) ---
        slot_l = self.hfss.modeler.get_variable_value("slotL")
        slot_w = self.hfss.modeler.get_variable_value("slotW")
        slot_sep = self.hfss.modeler.get_variable_value("slotSep")
        
        # Slots orientados ao longo do eixo Y, ortogonais ao dipolo
        slot1 = self.hfss.modeler.create_box(
            [cx - slot_w / 2.0, cy - slot_sep / 2.0 - slot_l / 2.0, z_gnd],
            [slot_w, slot_l, 0.001], name=f"{name_prefix}_Slot1_tool", matname="vacuum"
        )
        slot2 = self.hfss.modeler.create_box(
            [cx - slot_w / 2.0, cy + slot_sep / 2.0 - slot_l / 2.0, z_gnd],
            [slot_w, slot_l, 0.001], name=f"{name_prefix}_Slot2_tool", matname="vacuum"
        )
        
        # --- Radiating Patch ---
        patch_w_val = self.hfss.modeler.get_variable_value("patchW")
        patch_l_val = self.hfss.modeler.get_variable_value("patchL")
        patch = self.hfss.modeler.create_rectangle(
            "XY", [cx - patch_w_val / 2.0, cy - patch_l_val / 2.0, z_patch],
            ["patchW", "patchL"], name=f"{name_prefix}_Patch", matname="copper"
        )

        return [arm1, arm2], [slot1, slot2], patch

    # ---------- Pós-solve helpers ----------
    def _edit_sources_with_vars(self, excitations: List[str], magnitudes: List[str], phases: List[str]) -> bool:
        """Chama *Solutions.EditSources* ligando cada excitação a magnitude/fase (strings)."""
        try:
            sol = self.hfss.odesign.GetModule("Solutions")
            header = ["IncludePortPostProcessing:=", False, "SpecifySystemPower:=", False]
            cmd = [header]
            for ex, mag, ph in zip(excitations, magnitudes, phases):
                cmd.append(["Name:=", ex, "Magnitude:=", mag, "Phase:=", ph])
            sol.EditSources(cmd)
            self.log_message(f"Solutions.EditSources applied to {len(excitations)} port(s).")
            return True
        except Exception as e:
            self.log_message(f"EditSources failed: {e}")
            return False

    def _ensure_infinite_sphere(self, name="Infinite Sphere1") -> Optional[str]:
        """Cria (ou recria) *Infinite Sphere* com amostragem 1° x 1°, Theta 0→180."""
        try:
            rf = self.hfss.odesign.GetModule("RadField")
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
                     "ThetaStart:=", "0deg", "ThetaStop:=", "180deg", "ThetaStep:=", "1deg",
                     "PhiStart:=", "-180deg", "PhiStop:=", "180deg", "PhiStep:=", "1deg",
                     "UseLocalCS:=", False]
            rf.InsertInfiniteSphereSetup(props)
            self.log_message(f"Infinite sphere '{name}' created.")
            return name
        except Exception as e:
            self.log_message(f"Infinite sphere creation failed: {e}")
            return None

    def _postprocess_after_solve(self):
        """Configura fontes iniciais (1W/0deg) e constroi UI de beamforming."""
        try:
            exs = self._list_excitations()
            if not exs:
                self.log_message("No excitations found for post-processing.")
                return
            self._edit_sources_with_vars(exs, ["1W"] * len(exs), ["0deg"] * len(exs))
            self.populate_source_controls(exs)
            self.log_message("Initial sources set to 1W/0deg.")
        except Exception as e:
            self.log_message(f"Postprocess-after-solve error: {e}")

    # ------------- Helpers de solução -------------
    def _fetch_solution(self, expression: str, setup_candidates: Optional[List[str]] = None, **kwargs):
        """Wrapper robusto para `post.get_solution_data` tentando diferentes nomes de setup/sweep."""
        if setup_candidates is None:
            setup_candidates = ["Setup1 : LastAdaptive", "Setup1:LastAdaptive", "Setup1 : Sweep1", "Setup1:Sweep1"]
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
        """Converte retorno de `data_real()` em ndarray 1D com tamanho esperado."""
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
        """Retorna (ângulo, ganho_dB) para corte Theta ou Phi em `frequency` GHz."""
        try:
            expr = "dB(GainTotal)"
            if cut.lower() == "theta":
                variations = {"Freq": f"{frequency}GHz", "Theta": "All", "Phi": f"{fixed_angle_deg}deg"}
                prim = "Theta"
            else:
                variations = {"Freq": f"{frequency}GHz", "Theta": f"{fixed_angle_deg}deg", "Phi": "All"}
                prim = "Phi"
            sd = self._fetch_solution(
                expr,
                setup_candidates=["Setup1 : LastAdaptive", "Setup1:LastAdaptive", "Setup1 : Sweep1", "Setup1:Sweep1"],
                primary_sweep_variable=prim, variations=variations, context="Infinite Sphere1"
            )
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
                    setup_candidates=["Setup1 : LastAdaptive", "Setup1:LastAdaptive", "Setup1 : Sweep1", "Setup1:Sweep1"],
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
                        continue
                G_list.append(g)
            if TH_list is None or len(G_list) == 0:
                return None
            G = np.vstack(G_list).T  # shape (Ntheta, Nphi)
            PH = phi_vals[:G.shape[1]]
            TH = TH_list
            return TH, PH, G
        except Exception as e:
            self.log_message(f"3D grid error: {e}")
            return None
    # ------------- Simulação -------------
    def start_simulation_thread(self):
        """Inicia a simulação em *thread* separada para manter a UI responsiva."""
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
        """Fluxo completo: criar projeto, geometria, *setup/sweep*, analisar e pós-processar."""
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

            self._open_or_create_project()
            self.progress_bar.set(0.1)

            self.hfss.modeler.model_units = "mm"; self.log_message("Model units set to: mm")
            self._set_design_variables()
            self.created_ports.clear()

            # --- Garantir Materiais ---
            self._ensure_material("patch_mat", self.params["er_patch"], self.params["tan_d_patch"])
            self._ensure_material("feed_mat", self.params["er_feed"], self.params["tan_d_feed"])
            if self.params["honeycomb_material"] == "Nomex":
                 self._ensure_material("Nomex", 1.05, 0.001)

            # --- Construção do Stack-up ---
            self.log_message("Creating multi-layer stack-up")
            z_start = 0
            self.hfss.modeler.create_box(
                ["-subW/2", "-subL/2", z_start], ["subW", "subL", "h_feed_sub"], name="FeedSubstrate", matname="feed_mat"
            )
            z_start += self.hfss.modeler.get_variable_value("h_feed_sub")
            self.hfss.modeler.create_box(
                ["-subW/2", "-subL/2", z_start], ["subW", "subL", "h_air"], name="AirGap", matname="Air"
            )
            z_start += self.hfss.modeler.get_variable_value("h_air")
            ground = self.hfss.modeler.create_rectangle(
                "XY", ["-subW/2", "-subL/2", z_start], ["subW", "subL"], name="Ground", matname="copper"
            )
            self.hfss.modeler.create_box(
                ["-subW/2", "-subL/2", z_start], ["subW", "subL", "h_sub"], name="PatchSubstrate", matname="patch_mat"
            )
            z_start += self.hfss.modeler.get_variable_value("h_sub")
            self.hfss.modeler.create_box(
                ["-subW/2", "-subL/2", z_start], ["subW", "subL", "h_honey"], name="Honeycomb", matname=self.params["honeycomb_material"]
            )
            self.progress_bar.set(0.2)
            
            # --- Criação dos Elementos do Array ---
            self.log_message(f"Creating {self.calculated_params['num_patches']} slot-coupled elements")
            rows, cols = self.calculated_params["rows"], self.calculated_params["cols"]
            W, L = self.calculated_params["patch_width"], self.calculated_params["patch_length"]
            spacing = self.calculated_params["spacing"]
            total_w = cols * W + (cols - 1) * spacing
            total_l = rows * L + (rows - 1) * spacing
            start_x = -total_w / 2 + W / 2
            start_y = -total_l / 2 + L / 2
            
            all_slot_tools = []
            count = 0
            for r in range(rows):
                for c in range(cols):
                    if self.stop_simulation: self.log_message("Simulation stopped by user"); return
                    count += 1
                    cx = start_x + c * (W + spacing)
                    cy = start_y + r * (L + spacing)
                    self.log_message(f"Creating element {count} at ({r}, {c})")
                    _, slot_tools, _ = self._create_slotted_coupled_element(cx, cy, f"P{count}")
                    all_slot_tools.extend(slot_tools)
                    self.progress_bar.set(0.2 + 0.4 * (count / float(rows * cols)))
            
            # --- Operações Booleanas ---
            if all_slot_tools:
                self.log_message(f"Subtracting {len(all_slot_tools)} slots from ground plane")
                self.hfss.modeler.subtract(ground, all_slot_tools, keep_originals=False)

            self.hfss.assign_perfecte_to_sheets([o.name for o in self.hfss.modeler.get_objects_by_material("copper")])
            self.log_message("PerfectE assigned to all copper objects")
            
            # --- Contorno de Radiação ---
            self.log_message("Creating air region + radiation boundary")
            lambda0_mm = self.c / (self.params["sweep_start"] * 1e9) * 1000.0
            pad_mm = float(lambda0_mm) / 4.0
            region = self.hfss.modeler.create_region([pad_mm]*6, is_percentage=False)
            self.hfss.assign_radiation_boundary_to_objects(region)
            self.progress_bar.set(0.65)

            # --- Setup e Análise ---
            self._ensure_infinite_sphere("Infinite Sphere1")
            self.log_message("Creating simulation setup")
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = "f0"
            setup.props["MaxDeltaS"] = 0.02
            setup.props["SaveRadFieldsOnly"] = True

            self.log_message(f"Creating frequency sweep: {self.params['sweep_type']}")
            if self.params["sweep_type"] == "Discrete":
                setup.create_linear_step_sweep(unit="GHz", start_frequency=self.params["sweep_start"],
                                               stop_frequency=self.params["sweep_stop"], step_size=self.params["sweep_step"], name="Sweep1")
            else:
                setup.create_frequency_sweep(unit="GHz", name="Sweep1",
                                             start_frequency=self.params["sweep_start"],
                                             stop_frequency=self.params["sweep_stop"], sweep_type=self.params["sweep_type"])
            
            self.log_message(f"Validating design. Excitations created: {self._list_excitations()}")
            self.hfss.validate_full_design()
            self.log_message("Starting analysis")
            if self.save_project: self.hfss.save_project()
            
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])
            if self.stop_simulation: self.log_message("Simulation stopped by user"); return

            self._postprocess_after_solve()
            self.progress_bar.set(0.9)
            self.log_message("Processing results")
            self.analyze_and_mark_s11()
            self.refresh_patterns_only()
            self.progress_bar.set(1.0)
            self.sim_status_label.configure(text="Simulation completed")
            self.log_message("Simulation completed successfully")
        except Exception as e:
            self.log_message(f"Error in simulation: {e}\nTraceback: {traceback.format_exc()}")
            self.sim_status_label.configure(text=f"Simulation error: {e}")
        finally:
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.is_simulation_running = False

    # ------------- S11 / VSWR / |Z| -------------
    def _get_s11_curves(self):
        """Obtém f, S11(dB) e (opcional) re/im de S11 para cálculo de Z."""
        exs = self._list_excitations()
        if not exs:
            return None
        port_name = exs[0].split(":")[0]

        # tentar por nome (compatível com portas nomeadas)
        for expr_tpl in [f"( {port_name},{port_name} )", f"({port_name},{port_name})"]:
            sd_db = self._fetch_solution(
                f"dB(S{expr_tpl})",
                setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1", "Setup1 : LastAdaptive"]
            )
            if sd_db and hasattr(sd_db, "primary_sweep_values"):
                f = np.asarray(sd_db.primary_sweep_values, dtype=float)
                y_db = self._shape_series(sd_db.data_real(), f.size)
                if y_db.size == f.size and f.size > 0:
                    sd_re = self._fetch_solution(f"re(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
                    sd_im = self._fetch_solution(f"im(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
                    reS = self._shape_series(sd_re.data_real(), f.size) if sd_re else None
                    imS = self._shape_series(sd_im.data_real(), f.size) if sd_im else None
                    return f, y_db, reS, imS

        # fallback por índice
        expr_tpl = "(1,1)"
        sd_db = self._fetch_solution(
            f"dB(S{expr_tpl})",
            setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1", "Setup1 : LastAdaptive"]
        )
        if not sd_db:
            return None
        f = np.asarray(sd_db.primary_sweep_values, dtype=float)
        y_db = self._shape_series(sd_db.data_real(), f.size)
        sd_re = self._fetch_solution(f"re(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
        sd_im = self._fetch_solution(f"im(S{expr_tpl})", setup_candidates=["Setup1 : Sweep1", "Setup1:Sweep1"])
        reS = self._shape_series(sd_re.data_real(), f.size) if sd_re else None
        imS = self._shape_series(sd_im.data_real(), f.size) if sd_im else None
        return f, y_db, reS, imS

    def analyze_and_mark_s11(self):
        """Plota S11 e VSWR; estima Z=50*(1+S)/(1-S) no mínimo de S11, se possível."""
        try:
            self.ax_s11.clear(); self.ax_imp.clear()

            data = self._get_s11_curves()
            if not data:
                self.log_message("Solution Data failed to load. Check solution, context or expression.")
                self.canvas.draw(); return
            f, s11_db, reS, imS = data
            if f.size == 0 or s11_db.size == 0:
                self.log_message("S11 analysis aborted: empty curve.")
                self.canvas.draw(); return

            # S11 dB
            self.ax_s11.plot(f, s11_db, linewidth=2, label="S11 (dB)")
            self.ax_s11.axhline(y=-10, linestyle='--', alpha=0.7, label='-10 dB')
            self.ax_s11.set_xlabel("Frequency (GHz)")
            self.ax_s11.set_ylabel("S11 (dB)")
            self.ax_s11.set_title("S11 & VSWR")
            self.ax_s11.grid(True, alpha=0.5)

            # VSWR via |S|
            s_abs = 10 ** (s11_db / 20.0)
            s_abs = np.clip(s_abs, 0, 0.999999)
            vswr = (1 + s_abs) / (1 - s_abs)
            ax_v = self.ax_s11.twinx()
            ax_v.plot(f, vswr, linestyle='--', alpha=0.8, label='VSWR')
            ax_v.set_ylabel("VSWR")

            # mínimo de S11
            idx_min = int(np.argmin(s11_db))
            f_res = float(f[idx_min]); s11_min_db = float(s11_db[idx_min])
            self.ax_s11.scatter([f_res], [s11_min_db], s=45, marker="o", zorder=5)
            self.ax_s11.annotate(f"f_res={f_res:.4g} GHz\nS11={s11_min_db:.2f} dB",
                                 (f_res, s11_min_db), textcoords="offset points", xytext=(8, -16))
            cf = float(self.params["frequency"])
            self.ax_s11.axvline(x=cf, linestyle=':', alpha=0.7, color='r', label=f"f0={cf:g} GHz")

            # Impedância |Z|
            Zmag = None; R = X = None
            if reS is not None and imS is not None and reS.size == imS.size == f.size:
                S = reS + 1j*imS
                Z0 = 50.0
                with np.errstate(divide='ignore', invalid='ignore'):
                    Z = Z0 * (1 + S) / (1 - S)
                Zmag = np.abs(Z)
                self.ax_imp.plot(f, Zmag, linewidth=2)
                self.ax_imp.set_xlabel("Frequency (GHz)")
                self.ax_imp.set_ylabel("|Z| (Ω)")
                self.ax_imp.set_title("Input Impedance Magnitude")
                self.ax_imp.grid(True, alpha=0.5)
                Zr = Z[idx_min]
                R = float(np.real(Zr)); X = float(np.imag(Zr))

            self.last_s11_analysis = {"f": f, "s11_db": s11_db, "vswr": vswr,
                                      "Zmag": Zmag, "f_res": f_res,
                                      "R": R, "X": X}

            if R is not None and X is not None:
                self.result_label.configure(text=f"Min @ {f_res:.4g} GHz, S11={s11_min_db:.2f} dB, Z≈{R:.1f} + j{X:.1f} Ω")
            else:
                self.result_label.configure(text=f"Min @ {f_res:.4g} GHz, S11={s11_min_db:.2f} dB")

            self.canvas.draw()
        except Exception as e:
            self.log_message(f"Analyze S11 error: {e}\nTraceback: {traceback.format_exc()}")

    # ------------- Padrões / 3D -------------
    def refresh_patterns_only(self):
        """Atualiza cortes theta/phi e superfície 3D com base na solução atual."""
        try:
            self.ax_th.clear(); self.ax_ph.clear(); self.ax_3d.clear()

            f0 = float(self.params["frequency"])

            th, gth = self._get_gain_cut(f0, cut="theta", fixed_angle_deg=0.0)
            if th is not None and gth is not None:
                self.theta_cut = (th, gth)
                self.ax_th.plot(th, gth, linewidth=2)
                self.ax_th.set_xlabel("Theta (deg)"); self.ax_th.set_ylabel("Gain (dB)")
                self.ax_th.set_title("Radiation Pattern - Theta cut (Phi=0°)")
                self.ax_th.grid(True, alpha=0.5)
            else:
                self.ax_th.text(0.5, 0.5, "Theta-cut gain not available",
                                transform=self.ax_th.transAxes, ha="center", va="center")

            ph, gph = self._get_gain_cut(f0, cut="phi", fixed_angle_deg=90.0)
            if ph is not None and gph is not None:
                self.phi_cut = (ph, gph)
                self.ax_ph.plot(ph, gph, linewidth=2)
                self.ax_ph.set_xlabel("Phi (deg)"); self.ax_ph.set_ylabel("Gain (dB)")
                self.ax_ph.set_title("Radiation Pattern - Phi cut (Theta=90°)")
                self.ax_ph.grid(True, alpha=0.5)
            else:
                self.ax_ph.text(0.5, 0.5, "Phi-cut gain not available",
                                transform=self.ax_ph.transAxes, ha="center", va="center")

            # 3D
            grid = self._get_gain_3d_grid(f0, theta_step=self.params["theta_step"], phi_step=self.params["phi_step"])
            if grid is not None:
                TH_deg, PH_deg, Gdb = grid  # shapes (Nt, Np)
                self.grid3d = grid
                TH = np.deg2rad(TH_deg)[:, None] * np.ones((1, Gdb.shape[1]))
                PH = np.deg2rad(PH_deg)[None, :] * np.ones((Gdb.shape[0], 1))
                Glin = 10 ** (Gdb / 20.0)
                Glin = Glin - np.min(Glin)
                if np.max(Glin) > 0:
                    Glin = Glin / np.max(Glin)
                R = 0.2 + 0.8 * Glin
                X = R * np.sin(TH) * np.cos(PH)
                Y = R * np.sin(TH) * np.sin(PH)
                Z = R * np.cos(TH)
                self.ax_3d.plot_surface(X, Y, Z, rstride=1, cstride=1, linewidth=0, antialiased=True,
                                        cmap=cm.jet, shade=True)
                self.ax_3d.set_title("3D Gain Pattern (normalized)")
                self.ax_3d.set_axis_off()
            else:
                self.ax_3d.text2D(0.5, 0.5, "3D pattern not available", transform=self.ax_3d.transAxes, ha="center", va="center")

            self.fig.tight_layout()
            self.canvas.draw()
            self.log_message("Patterns refreshed.")
        except Exception as e:
            self.log_message(f"Refresh patterns error: {e}\nTraceback: {traceback.format_exc()}")

    # ------------- Beamforming UI -------------
    def populate_source_controls(self, excitations: List[str]):
        """(Re)constrói controles de potência/fase por porta para *beamforming*."""
        for child in self.src_frame.winfo_children():
            child.destroy()
        self.source_controls.clear()

        if not excitations:
            ctk.CTkLabel(self.src_frame, text="No excitations found.").pack(padx=8, pady=6)
            return

        head = ctk.CTkFrame(self.src_frame); head.pack(fill="x", padx=8, pady=4)
        ctk.CTkLabel(head, text="Beamforming & Refresh", font=ctk.CTkFont(weight="bold")).pack(side="left")
        grid = ctk.CTkFrame(self.src_frame); grid.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(grid, text="Port", width=120).grid(row=0, column=0, padx=4, pady=4, sticky="w")
        ctk.CTkLabel(grid, text="Power (W)", width=120).grid(row=0, column=1, padx=4, pady=4)
        ctk.CTkLabel(grid, text="Phase (deg)", width=300).grid(row=0, column=2, padx=4, pady=4)

        for i, ex in enumerate(excitations, start=1):
            row = i
            ctk.CTkLabel(grid, text=ex.split(":")[0], width=120).grid(row=row, column=0, padx=4, pady=3, sticky="w")
            p_entry = ctk.CTkEntry(grid, width=100); p_entry.insert(0, "1.0")
            p_entry.grid(row=row, column=1, padx=4, pady=3)
            phase_slider = ctk.CTkSlider(grid, from_=0, to=360, number_of_steps=360)
            phase_slider.set(0); phase_slider.grid(row=row, column=2, padx=6, pady=6, sticky="ew")
            grid.grid_columnconfigure(2, weight=1)
            self.source_controls[ex] = {"power": p_entry, "phase": phase_slider}

    def apply_sources_from_ui(self):
        """Lê controles de UI e redefine fontes via EditSources; atualiza padrões."""
        try:
            exs = self._list_excitations()
            if not exs:
                self.log_message("No excitations to apply.")
                return
            pvals, phvals = [], []
            for i, ex in enumerate(exs, start=1):
                ctrl = self.source_controls.get(ex)
                if not ctrl:
                    continue
                try:
                    pw = float(ctrl["power"].get())
                except Exception:
                    pw = 1.0
                ph = float(ctrl["phase"].get())
                pvals.append(f"{pw}W")
                phvals.append(f"{ph}deg")
            self._edit_sources_with_vars(exs, pvals, phvals)
            self.refresh_patterns_only()
        except Exception as e:
            self.log_message(f"Apply sources error: {e}\nTraceback: {traceback.format_exc()}")

    def toggle_auto_refresh(self):
        if self.auto_refresh_var.get():
            self.schedule_auto_refresh()
        else:
            if self.auto_refresh_job:
                try:
                    self.window.after_cancel(self.auto_refresh_job)
                except Exception:
                    pass
                self.auto_refresh_job = None

    def schedule_auto_refresh(self):
        self.apply_sources_from_ui() # Apply sources and then refresh patterns
        if self.auto_refresh_var.get():
            self.auto_refresh_job = self.window.after(1500, self.schedule_auto_refresh)

    # ------------- Exportações -------------
    def export_csv(self):
        """Exporta S11 (f, dB) para CSV."""
        try:
            if self.last_s11_analysis is not None:
                f = self.last_s11_analysis["f"]; s11 = self.last_s11_analysis["s11_db"]
                np.savetxt("simulation_results.csv", np.column_stack((f, s11)),
                           delimiter=",", header="Frequency (GHz), S11 (dB)", comments='')
                self.log_message("Data exported to simulation_results.csv")
            else:
                self.log_message("Run 'Analyze S11' first.")
        except Exception as e:
            self.log_message(f"Error exporting CSV: {e}")

    def export_png(self):
        """Exporta *figure* atual para PNG de alta resolução."""
        try:
            if hasattr(self, 'fig'):
                self.fig.savefig("simulation_results.png", dpi=300, bbox_inches='tight')
                self.log_message("Plot saved to simulation_results.png")
        except Exception as e:
            self.log_message(f"Error saving plot: {e}")

    # ------------- Cleanup / Persistência -------------
    def cleanup(self):
        """Fecha projeto/desktop e limpa pasta temporária (se não salvar)."""
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
                    self.desktop.release_desktop(close_projects=True, close_on_exit=True)
                except Exception as e:
                    self.log_message(f"Error releasing desktop: {e}")
            if self.temp_folder and not self.save_project:
                try:
                    self.temp_folder.cleanup()
                except Exception as e:
                    self.log_message(f"Error cleaning temp files: {e}")
        except Exception as e:
            self.log_message(f"Error during cleanup: {e}")

    def on_closing(self):
        self.log_message("Application closing...")
        self.cleanup()
        self.window.quit(); self.window.destroy()

    def save_parameters(self):
        try:
            all_params = {**self.params, **self.calculated_params}
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
            self.update_interface_from_params()
            self.log_message("Parameters loaded from antenna_parameters.json")
        except Exception as e:
            self.log_message(f"Error loading parameters: {e}")

    def update_interface_from_params(self):
        try:
            for key, widget in self.entries:
                if key in self.params:
                    value = self.params[key]
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end"); widget.insert(0, str(value))
                    elif isinstance(widget, ctk.StringVar):
                        widget.set(str(value))
                elif key == "show_gui" and "non_graphical" in self.params:
                    widget.set(not self.params["non_graphical"])
                elif key == "save_project":
                    widget.set(self.save_project)

            # Update calculated labels
            self.calculate_parameters()
            self.log_message("Interface updated with loaded parameters")
        except Exception as e:
            self.log_message(f"Error updating interface: {e}")

    def run(self):
        try:
            self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.window.mainloop()
        finally:
            self.cleanup()


if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()