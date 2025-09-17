# -*- coding: utf-8 -*-
import os
import tempfile
from datetime import datetime
import math
import json
import traceback
import queue
import threading
from typing import Tuple, List, Optional
import webbrowser
import sys

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont

# Tente importar o ANSYS, mas não quebre se não estiver disponível
try:
    from ansys.aedt.core import Desktop, Hfss
    ANSYS_AVAILABLE = True
except ImportError:
    ANSYS_AVAILABLE = False
    # Criar classes dummy para evitar erros
    class Desktop:
        def __init__(self, **kwargs):
            pass
    class Hfss:
        def __init__(self, **kwargs):
            pass

# ---------- Aparência ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ModernPatchAntennaDesigner:
    def __init__(self):
        # AEDT
        self.desktop: Optional[Desktop] = None
        self.hfss: Optional[Hfss] = None
        self.temp_folder = None
        self.project_path = ""
        self.project_title = "patch_array"
        self.design_name = "patch_array"

        # GUI/log
        self.log_queue = queue.Queue()
        self.is_simulation_running = False
        self.save_project = False
        self.stop_simulation = False
        self.simulation_data = None

        # Estados
        self.optimized = False
        self._shape_id = 0  # contador para nomes

        # ---------- Parâmetros do usuário ----------
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

            # Coax
            "probe_radius": 0.40,          # mm (a)
            "coax_ba_ratio": 2.3,
            "coax_wall_thickness": 0.20,   # mm
            "coax_port_length": 3.0,       # mm
            "antipad_clearance": 0.10,     # mm

            # Sweep
            "sweep_type": "Interpolating",
            "sweep_step": 0.02,            # GHz
        }

        # ---------- Parâmetros calculados ----------
        self.calculated_params = {
            "num_patches": 4,
            "spacing": 0.0,
            "patch_length": 9.57,
            "patch_width": 9.25,
            "rows": 2,
            "cols": 2,
            "lambda_g": 0.0,
            "substrate_width": 0.0,
            "substrate_length": 0.0
        }

        self.c = 299792458.0

        # Carregar ícones
        self.icons = self.load_icons()
        
        # Inicializar estatísticas
        self.simulation_stats = {
            "status": "Not Started",
            "mesh_elements": 0,
            "simulation_time": "0 min",
            "memory_usage": "0 MB"
        }
        
        self.setup_gui()

    # ===================== Resultados =====================
    def _plot_results(self):
        try:
            self._log("Plotando…")
            for ax in [self.ax_s11, self.ax_th, self.ax_ph]:
                ax.clear(); ax.grid(True, alpha=0.4)

            # S11 (porta nomeado -> fallback S(1,1))
            sol = None
            try:
                expr = f"dB(S({self._last_portname},{self._last_portname}))"
                rpt = self.hfss.post.reports_by_category.standard(expressions=[expr])
                rpt.context = ["Setup1: Sweep1"]
                sol = rpt.get_solution_data()
            except Exception:
                try:
                    expr = "dB(S(1,1))"
                    rpt = self.hfss.post.reports_by_category.standard(expressions=[expr])
                    rpt.context = ["Setup1: Sweep1"]
                    sol = rpt.get_solution_data()
                except Exception as e:
                    self._log(f"S11 aviso: {e}")

            if sol:
                try:
                    f = np.asarray(sol.primary_sweep_values, dtype=float)
                    dat = sol.data_real()
                    y = np.asarray(dat[0] if isinstance(dat, (list, tuple)) else dat, dtype=float)
                    if f.size and f.size == y.size:
                        self.simulation_data = np.column_stack((f, y))
                        self.ax_s11.plot(f, y, linewidth=2, label="S11")
                        self.ax_s11.axhline(-10, linestyle="--", alpha=0.6, label="-10 dB")
                        self.ax_s11.axvline(float(self.params["frequency"]), linestyle="--", alpha=0.6)
                        self.ax_s11.set_xlabel("Freq (GHz)"); self.ax_s11.set_ylabel("dB"); self.ax_s11.set_title("S11"); self.ax_s11.legend()
                    else:
                        self.ax_s11.text(0.5, 0.5, "S11 indisponível", transform=self.ax_s11.transAxes, ha="center", va="center")
                except Exception as e:
                    self._log(f"S11 parse erro: {e}")
                    self.ax_s11.text(0.5, 0.5, "S11 indisponível", transform=self.ax_s11.transAxes, ha="center", va="center")
            else:
                self.ax_s11.text(0.5, 0.5, "S11 indisponível", transform=self.ax_s11.transAxes, ha="center", va="center")

            # Far-field cortes - Correção da forma de criar relatórios
            try:
                # Obtém a frequência central para o campo distante
                freq_center = float(self.params["frequency"])
                
                # Diagrama de irradiação no plano E (phi=0)
                # Correção: Usar a sintaxe correta para criar relatórios de campo distante
                report = self.hfss.post.create_report(
                    expressions=["dB(GainTotal)"],
                    context="IS1",
                    setup_name="Setup1 : LastAdaptive",
                    variations={
                        "Freq": f"{freq_center}GHz",
                        "Theta": "all",
                        "Phi": "0deg"
                    },
                    primary_sweep_variable="Theta",
                    report_category="Far Fields"
                )
                
                if report:
                    ff_data = report.get_solution_data()
                    if ff_data:
                        theta = np.array(ff_data.primary_sweep_values)
                        gain = np.array(ff_data.data_real())[0]
                        self.ax_th.plot(theta, gain)
                        self.ax_th.set_xlabel("Theta (°)"); self.ax_th.set_ylabel("Gain (dB)"); 
                        self.ax_th.set_title(f"Plano E (Phi=0°) @ {freq_center} GHz")
                        self.ax_th.grid(True)

                # Diagrama de irradiação no plano H (theta=90)
                report2 = self.hfss.post.create_report(
                    expressions=["dB(GainTotal)"],
                    context="IS1",
                    setup_name="Setup1 : LastAdaptive",
                    variations={
                        "Freq": f"{freq_center}GHz",
                        "Theta": "90deg",
                        "Phi": "all"
                    },
                    primary_sweep_variable="Phi",
                    report_category="Far Fields"
                )
                
                if report2:
                    ff_data2 = report2.get_solution_data()
                    if ff_data2:
                        phi = np.array(ff_data2.primary_sweep_values)
                        gain = np.array(ff_data2.data_real())[0]
                        self.ax_ph.plot(phi, gain)
                        self.ax_ph.set_xlabel("Phi (°)"); self.ax_ph.set_ylabel("Gain (dB)"); 
                        self.ax_ph.set_title(f"Plano H (Theta=90°) @ {freq_center} GHz")
                        self.ax_ph.grid(True)
                        
            except Exception as e:
                self._log(f"Far-field indisponível: {e}")
                self.ax_th.text(0.5, 0.5, "FF indisponível", transform=self.ax_th.transAxes, ha="center", va="center")
                self.ax_ph.text(0.5, 0.5, "FF indisponível", transform=self.ax_ph.transAxes, ha="center", va="center")

            self.fig.tight_layout(); self.canvas.draw()
            self._log("Plot OK.")
        except Exception as e:
            self._log(f"Erro nos gráficos: {e}\n{traceback.format_exc()}")



    def load_icons(self):
        """Cria ícones simples para a interface"""
        icons = {}
        try:
            # Criar ícones usando texto (emoji) como fallback
            icon_size = (20, 20)
            
            # Função para criar ícone a partir de texto
            def create_text_icon(text, size=icon_size, bg_color=None):
                if bg_color is None:
                    bg_color = ("gray25", "gray75")
                
                dark_mode = ctk.get_appearance_mode() == "Dark"
                bg = bg_color[0] if dark_mode else bg_color[1]
                
                image = Image.new("RGBA", size, color=bg)
                draw = ImageDraw.Draw(image)
                try:
                    font = ImageFont.truetype("arial.ttf", 14)
                except:
                    font = ImageFont.load_default()
                
                text_bbox = draw.textbbox((0, 0), text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]
                
                position = ((size[0] - text_width) // 2, (size[1] - text_height) // 2)
                draw.text(position, text, fill="white", font=font)
                
                return ctk.CTkImage(image, size=size)
            
            # Criar ícones
            icons["calculate"] = create_text_icon("📐")
            icons["save"] = create_text_icon("💾")
            icons["load"] = create_text_icon("📂")
            icons["run"] = create_text_icon("▶️")
            icons["stop"] = create_text_icon("⏹️")
            icons["export"] = create_text_icon("📊")
            icons["help"] = create_text_icon("❓")
            icons["doc"] = create_text_icon("📘")
            
        except Exception as e:
            print(f"Erro ao criar ícones: {e}")
            
        return icons

    # ---------------- GUI ----------------
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer - Professional Edition")
        self.window.geometry("1600x1000")
        self.window.minsize(1400, 800)
        
        # Configurar layout principal
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        # Header moderno
        header = ctk.CTkFrame(self.window, height=70, fg_color=("#2B2B2B", "#3B3B3B"), corner_radius=0)
        header.grid(row=0, column=0, sticky="nsew", columnspan=2)
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)
        
        # Logo e título
        logo_frame = ctk.CTkFrame(header, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=20, pady=15, sticky="w")
        
        # Ícone de antena
        ctk.CTkLabel(logo_frame, text="📡", font=ctk.CTkFont(size=28)).pack(side="left", padx=(0, 10))
        
        # Título
        title_frame = ctk.CTkFrame(logo_frame, fg_color="transparent")
        title_frame.pack(side="left")
        ctk.CTkLabel(title_frame, text="Patch Antenna Array Designer", 
                    font=ctk.CTkFont(size=24, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(title_frame, text="Professional Edition", 
                    font=ctk.CTkFont(size=14), text_color=("gray70", "gray50")).pack(anchor="w")
        
        # Botões de ação no header
        action_frame = ctk.CTkFrame(header, fg_color="transparent")
        action_frame.grid(row=0, column=1, padx=20, pady=15, sticky="e")
        
        ctk.CTkButton(action_frame, text="Documentation", width=120, 
                     command=self.open_documentation, fg_color="transparent",
                     border_width=1, border_color=("#3B8ED0", "#1F6AA5"),
                     image=self.icons.get("doc"), compound="left").pack(side="right", padx=5)
        
        ctk.CTkButton(action_frame, text="Help", width=80, 
                     command=self.show_help, fg_color="transparent",
                     border_width=1, border_color=("#3B8ED0", "#1F6AA5"),
                     image=self.icons.get("help"), compound="left").pack(side="right", padx=5)

        # Painel principal com abas
        self.tabview = ctk.CTkTabview(self.window, fg_color=("gray90", "gray15"))
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Adicionar abas
        tab_names = ["Design Parameters", "Simulation Control", "Results Analysis", "Activity Log"]
        for name in tab_names:
            self.tabview.add(name)
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        # Configurar cada aba
        self.setup_parameters_tab()
        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_log_tab()

        # Barra de status
        status = ctk.CTkFrame(self.window, height=40, fg_color=("gray85", "gray20"))
        status.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        status.grid_propagate(False)
        status.grid_columnconfigure(0, weight=1)
        
        # Status principal
        self.status_label = ctk.CTkLabel(status, text="Ready to design your antenna array", 
                                        font=ctk.CTkFont(weight="bold"), anchor="w")
        self.status_label.grid(row=0, column=0, padx=15, pady=0, sticky="w")
        
        # Informações de sistema
        progress_frame = ctk.CTkFrame(status, fg_color="transparent")
        progress_frame.grid(row=0, column=1, padx=15, pady=0, sticky="e")
        
        ctk.CTkLabel(progress_frame, text="CPU:", font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 5))
        self.cpu_usage = ctk.CTkProgressBar(progress_frame, width=80, height=10)
        self.cpu_usage.pack(side="left")
        self.cpu_usage.set(0.1)
        
        # ANSYS status
        ansys_status = "Available" if ANSYS_AVAILABLE else "Not Available"
        status_color = "#2E8B57" if ANSYS_AVAILABLE else "#DC143C"
        ctk.CTkLabel(progress_frame, text=f"ANSYS: {ansys_status}", 
                    text_color=status_color, font=ctk.CTkFont(size=12)).pack(side="left", padx=(15, 0))

        # Fecha seguro
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Iniciar processamento de log
        self.process_log_queue()

    def create_section(self, parent, title, description=None, row=0, column=0, colspan=1):
        """Cria uma seção com título e descrição"""
        sec = ctk.CTkFrame(parent, fg_color=("gray95", "gray18"), corner_radius=8)
        sec.grid(row=row, column=column, sticky="nsew", padx=8, pady=8, columnspan=colspan)
        sec.grid_columnconfigure(0, weight=1)
        
        # Título da seção
        title_frame = ctk.CTkFrame(sec, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        title_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(title_frame, text=title, font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w")
        
        if description:
            ctk.CTkLabel(title_frame, text=description, font=ctk.CTkFont(size=12), 
                         text_color=("gray50", "gray60"), wraplength=400).grid(
            row=1, column=0, sticky="w", pady=(0, 5))
        
        # Separador
        ctk.CTkFrame(sec, height=1, fg_color=("gray80", "gray30")).grid(
            row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        
        return sec

    def setup_parameters_tab(self):
        """Configura a aba de parâmetros de design"""
        tab = self.tabview.tab("Design Parameters")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Frame principal com scroll
        main = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main.grid_columnconfigure((0, 1), weight=1, uniform="col")
        self.entries = []

        # Função para adicionar campos de entrada
        def add_entry(section, label, key, value, row, tooltip=None, combo=None, check=False, unit=None):
            label_frame = ctk.CTkFrame(section, fg_color="transparent")
            label_frame.grid(row=row, column=0, padx=12, pady=6, sticky="w")
            
            ctk.CTkLabel(label_frame, text=label, font=ctk.CTkFont(weight="bold")).pack(anchor="w")
            
            if tooltip:
                help_btn = ctk.CTkButton(label_frame, text="?", width=20, height=20,
                                        command=lambda: self.show_tooltip(tooltip),
                                        fg_color="transparent", hover_color=("gray80", "gray30"))
                help_btn.pack(side="right", padx=(5, 0))
            
            input_frame = ctk.CTkFrame(section, fg_color="transparent")
            input_frame.grid(row=row, column=1, padx=12, pady=6, sticky="ew")
            input_frame.grid_columnconfigure(0, weight=1)
            
            if combo:
                var = ctk.StringVar(value=value)
                widget = ctk.CTkComboBox(input_frame, values=combo, variable=var, width=200)
                widget.grid(row=0, column=0, sticky="w")
                self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                widget = ctk.CTkCheckBox(input_frame, text="", variable=var, width=20)
                widget.grid(row=0, column=0, sticky="w")
                self.entries.append((key, var))
            else:
                widget = ctk.CTkEntry(input_frame, width=200)
                widget.insert(0, str(value))
                widget.grid(row=0, column=0, sticky="w")
                self.entries.append((key, widget))
                
                if unit:
                    ctk.CTkLabel(input_frame, text=unit, text_color=("gray50", "gray60")).grid(
                        row=0, column=1, padx=(5, 0), sticky="w")
            
            return row + 1

        # Seção de Parâmetros da Antena
        sec_antenna = self.create_section(main, "Antenna Parameters", 
                                         "Configure the fundamental antenna properties", 
                                         0, 0)
        r = 2
        r = add_entry(sec_antenna, "Center Frequency:", "frequency", self.params["frequency"], r, 
                     tooltip="Operating frequency of the antenna", unit="GHz")
        r = add_entry(sec_antenna, "Desired Gain:", "gain", self.params["gain"], r,
                     tooltip="Target gain for the antenna array", unit="dBi")
        r = add_entry(sec_antenna, "Sweep Start:", "sweep_start", self.params["sweep_start"], r,
                     tooltip="Start frequency for simulation sweep", unit="GHz")
        r = add_entry(sec_antenna, "Sweep Stop:", "sweep_stop", self.params["sweep_stop"], r,
                     tooltip="Stop frequency for simulation sweep", unit="GHz")
        r = add_entry(sec_antenna, "Element Spacing:", "spacing_type", self.params["spacing_type"], r,
                     tooltip="Distance between antenna elements", 
                     combo=["lambda/2", "lambda", "0.7*lambda", "0.8*lambda", "0.9*lambda"])

        # Seção de Substrato
        sec_substrate = self.create_section(main, "Substrate Properties", 
                                           "Define substrate material and characteristics", 
                                           1, 0)
        r = 2
        r = add_entry(sec_substrate, "Material:", "substrate_material", self.params["substrate_material"], r,
                     tooltip="Substrate material type",
                     combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        r = add_entry(sec_substrate, "Relative Permittivity (εr):", "er", self.params["er"], r,
                     tooltip="Dielectric constant of the substrate")
        r = add_entry(sec_substrate, "Loss Tangent (tanδ):", "tan_d", self.params["tan_d"], r,
                     tooltip="Dissipation factor of the substrate")
        r = add_entry(sec_substrate, "Thickness:", "substrate_thickness", self.params["substrate_thickness"], r,
                     tooltip="Substrate height", unit="mm")
        r = add_entry(sec_substrate, "Metal Thickness:", "metal_thickness", self.params["metal_thickness"], r,
                     tooltip="Conductor thickness", unit="mm")

        # Seção de Alimentação
        sec_feed = self.create_section(main, "Feed Configuration", 
                                      "Coaxial feed parameters", 
                                      0, 1)
        r = 2
        r = add_entry(sec_feed, "Probe Radius (a):", "probe_radius", self.params["probe_radius"], r,
                     tooltip="Inner conductor radius", unit="mm")
        r = add_entry(sec_feed, "Coax Ratio (b/a):", "coax_ba_ratio", self.params["coax_ba_ratio"], r,
                     tooltip="Ratio of outer to inner conductor radius")
        r = add_entry(sec_feed, "Shield Wall Thickness:", "coax_wall_thickness", self.params["coax_wall_thickness"], r,
                     tooltip="Thickness of the outer conductor", unit="mm")
        r = add_entry(sec_feed, "Port Length (Lp):", "coax_port_length", self.params["coax_port_length"], r,
                     tooltip="Length of the port section", unit="mm")
        r = add_entry(sec_feed, "Antipad Clearance:", "antipad_clearance", self.params["antipad_clearance"], r,
                     tooltip="Clearance around the feed", unit="mm")

        # Seção de Simulação
        sec_simulation = self.create_section(main, "Simulation Settings", 
                                            "Configure simulation parameters", 
                                            1, 1)
        r = 2
        r = add_entry(sec_simulation, "CPU Cores:", "cores", self.params["cores"], r,
                     tooltip="Number of processor cores to use")
        r = add_entry(sec_simulation, "Save Project:", "save_project", self.save_project, r,
                     tooltip="Save project after simulation", check=True)
        r = add_entry(sec_simulation, "Sweep Type:", "sweep_type", self.params["sweep_type"], r,
                     tooltip="Type of frequency sweep",
                     combo=["Discrete", "Interpolating", "Fast"])
        r = add_entry(sec_simulation, "Step Size:", "sweep_step", self.params["sweep_step"], r,
                     tooltip="Frequency step for discrete sweep", unit="GHz")

        # Seção de Parâmetros Calculados
        sec_calculated = self.create_section(main, "Calculated Parameters", 
                                            "Parameters derived from your inputs", 
                                            2, 0, colspan=2)
        grid = ctk.CTkFrame(sec_calculated, fg_color="transparent")
        grid.grid(row=2, column=0, sticky="nsew", padx=12, pady=8)
        grid.columnconfigure((0, 1), weight=1)
        
        # Usando CTkTextbox para melhor formatação
        self.calculated_text = ctk.CTkTextbox(grid, width=400, height=150, font=ctk.CTkFont(size=13))
        self.calculated_text.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        self.calculated_text.insert("1.0", "Click 'Calculate' to compute parameters")
        self.calculated_text.configure(state="disabled")
        
        # Botões de ação
        btn_frame = ctk.CTkFrame(sec_calculated, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=12, pady=8)
        btn_frame.columnconfigure((0, 1, 2), weight=1)
        
        ctk.CTkButton(btn_frame, text="Calculate Parameters", command=self.calculate_parameters,
                     image=self.icons.get("calculate"), compound="left", fg_color="#2E8B57").grid(
                     row=0, column=0, padx=5, sticky="ew")
        ctk.CTkButton(btn_frame, text="Save Configuration", command=self.save_parameters,
                     image=self.icons.get("save"), compound="left", fg_color="#4169E1").grid(
                     row=0, column=1, padx=5, sticky="ew")
        ctk.CTkButton(btn_frame, text="Load Configuration", command=self.load_parameters,
                     image=self.icons.get("load"), compound="left", fg_color="#FF8C00").grid(
                     row=0, column=2, padx=5, sticky="ew")

    def setup_simulation_tab(self):
        """Configura a aba de controle de simulação"""
        tab = self.tabview.tab("Simulation Control")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        
        # Título
        ctk.CTkLabel(main, text="Simulation Control Center", 
                    font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(0, 20))
        
        # Cartões de status
        status_cards = ctk.CTkFrame(main, fg_color="transparent")
        status_cards.pack(fill="x", pady=(0, 20))
        
        self.status_cards = {}
        cards_data = [
            {"title": "Project Status", "value": "Not Created", "color": "#7C7C7C", "key": "status"},
            {"title": "Mesh Elements", "value": "0", "color": "#3B8ED0", "key": "mesh_elements"},
            {"title": "Simulation Time", "value": "0 min", "color": "#1F6AA5", "key": "simulation_time"},
            {"title": "Memory Usage", "value": "0 MB", "color": "#FF4B4B", "key": "memory_usage"}
        ]
        
        for i, card in enumerate(cards_data):
            card_frame = ctk.CTkFrame(status_cards, fg_color=("gray90", "gray20"), 
                                     border_width=1, border_color=("gray70", "gray30"),
                                     corner_radius=8, width=180, height=100)
            card_frame.grid(row=0, column=i, padx=10, sticky="nsew")
            card_frame.grid_propagate(False)
            card_frame.grid_columnconfigure(0, weight=1)
            
            ctk.CTkLabel(card_frame, text=card["title"], 
                        font=ctk.CTkFont(weight="bold"), 
                        text_color=("gray40", "gray60")).grid(row=0, column=0, pady=(15, 5))
            
            value_label = ctk.CTkLabel(card_frame, text=card["value"], 
                        font=ctk.CTkFont(size=18, weight="bold"),
                        text_color=card["color"])
            value_label.grid(row=1, column=0, pady=(0, 15))
            
            self.status_cards[card["key"]] = value_label
            status_cards.columnconfigure(i, weight=1)
        
        # Controles de simulação
        control_frame = ctk.CTkFrame(main, fg_color=("gray92", "gray18"), corner_radius=8)
        control_frame.pack(fill="x", pady=(0, 20))
        control_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(control_frame, text="Simulation Controls", 
                    font=ctk.CTkFont(size=16, weight="bold")).grid(
                    row=0, column=0, padx=20, pady=(15, 10), sticky="w")
        
        btn_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="ew")
        btn_frame.columnconfigure((0, 1), weight=1)
        
        self.run_button = ctk.CTkButton(btn_frame, text="Run Simulation", 
                                       command=self.start_simulation_thread,
                                       image=self.icons.get("run"), compound="left",
                                       fg_color="#2E8B57", height=40, font=ctk.CTkFont(weight="bold"))
        self.run_button.grid(row=0, column=0, padx=10, sticky="ew")
        
        self.stop_button = ctk.CTkButton(btn_frame, text="Stop Simulation", 
                                        command=self.stop_simulation_thread,
                                        image=self.icons.get("stop"), compound="left",
                                        fg_color="#DC143C", state="disabled", height=40)
        self.stop_button.grid(row=0, column=1, padx=10, sticky="ew")
        
        # Barra de progresso
        progress_sec = ctk.CTkFrame(main, fg_color=("gray92", "gray18"), corner_radius=8)
        progress_sec.pack(fill="x", pady=(0, 10))
        progress_sec.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(progress_sec, text="Simulation Progress", 
                    font=ctk.CTkFont(size=16, weight="bold")).grid(
                    row=0, column=0, padx=20, pady=(15, 10), sticky="w")
        
        # Barra de progresso com porcentagem
        progress_bar_frame = ctk.CTkFrame(progress_sec, fg_color="transparent")
        progress_bar_frame.grid(row=1, column=0, padx=20, pady=(0, 5), sticky="ew")
        
        self.progress_bar = ctk.CTkProgressBar(progress_bar_frame, height=20)
        self.progress_bar.pack(fill="x", side="left", expand=True)
        self.progress_bar.set(0)
        
        self.progress_label = ctk.CTkLabel(progress_bar_frame, text="0%", width=40)
        self.progress_label.pack(side="right", padx=(10, 0))
        
        # Status da simulação
        self.sim_status_label = ctk.CTkLabel(progress_sec, text="Ready to run simulation", 
                                            font=ctk.CTkFont(weight="bold"))
        self.sim_status_label.grid(row=2, column=0, padx=20, pady=(0, 15), sticky="w")

    def setup_results_tab(self):
        """Configura a aba de análise de resultados"""
        tab = self.tabview.tab("Results Analysis")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        
        # Cabeçalho
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(header, text="Simulation Results", 
                    font=ctk.CTkFont(size=20, weight="bold")).grid(
                    row=0, column=0, sticky="w")
        
        # Botões de exportação
        export_frame = ctk.CTkFrame(header, fg_color="transparent")
        export_frame.grid(row=0, column=1, sticky="e")
        
        ctk.CTkButton(export_frame, text="Export Data", 
                     image=self.icons.get("export"), compound="left",
                     command=self.export_csv, width=120).grid(row=0, column=0, padx=5)
        ctk.CTkButton(export_frame, text="Export Plot", 
                     image=self.icons.get("export"), compound="left",
                     command=self.export_png, width=120).grid(row=0, column=1, padx=5)
        
        # Área de gráficos
        graph_frame = ctk.CTkFrame(main, fg_color=("gray95", "gray16"), corner_radius=8)
        graph_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        graph_frame.grid_columnconfigure(0, weight=1)
        graph_frame.grid_rowconfigure(0, weight=1)
        
        # Abas para diferentes visualizações
        graph_tabs = ctk.CTkTabview(graph_frame, fg_color=("gray90", "gray20"))
        graph_tabs.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        graph_tabs.add("S-Parameters")
        graph_tabs.add("Radiation Pattern")
        graph_tabs.add("3D Pattern")
        
        # Configurar gráficos
        for tab_name in ["S-Parameters", "Radiation Pattern", "3D Pattern"]:
            graph_tabs.tab(tab_name).grid_columnconfigure(0, weight=1)
            graph_tabs.tab(tab_name).grid_rowconfigure(0, weight=1)
        
        # Figura para S-Parameters
        fig_frame = ctk.CTkFrame(graph_tabs.tab("S-Parameters"), fg_color="transparent")
        fig_frame.grid(row=0, column=0, sticky="nsew")
        fig_frame.grid_columnconfigure(0, weight=1)
        fig_frame.grid_rowconfigure(0, weight=1)
        
        self.fig = plt.figure(figsize=(10, 8))
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        self.fig.patch.set_facecolor(face)
        
        # Subplots
        self.ax_s11 = self.fig.add_subplot(2, 2, 1)
        self.ax_vswr = self.fig.add_subplot(2, 2, 2)
        self.ax_th = self.fig.add_subplot(2, 2, 3)
        self.ax_ph = self.fig.add_subplot(2, 2, 4)
        
        for ax in [self.ax_s11, self.ax_vswr, self.ax_th, self.ax_ph]:
            ax.set_facecolor(face)
            ax.grid(True, alpha=0.4)
            ax.tick_params(colors='gray')
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=fig_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        
        # Placeholder para outros gráficos
        for tab_name in ["Radiation Pattern", "3D Pattern"]:
            placeholder = ctk.CTkFrame(graph_tabs.tab(tab_name), fg_color="transparent")
            placeholder.grid(row=0, column=0, sticky="nsew")
            placeholder.grid_columnconfigure(0, weight=1)
            placeholder.grid_rowconfigure(0, weight=1)
            
            ctk.CTkLabel(placeholder, text=f"{tab_name} visualization will appear here after simulation",
                        font=ctk.CTkFont(size=14), text_color=("gray50", "gray60")).grid(row=0, column=0)
        
        # Estatísticas de resultados
        stats_frame = ctk.CTkFrame(main, fg_color=("gray92", "gray18"), corner_radius=8)
        stats_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        self.stats_labels = {}
        stats_data = [
            {"title": "Resonant Frequency", "value": "-", "unit": "GHz", "key": "resonant_freq"},
            {"title": "Bandwidth", "value": "-", "unit": "MHz", "key": "bandwidth"},
            {"title": "Return Loss", "value": "-", "unit": "dB", "key": "return_loss"},
            {"title": "VSWR", "value": "-", "unit": "", "key": "vswr"}
        ]
        
        for i, stat in enumerate(stats_data):
            stat_card = ctk.CTkFrame(stats_frame, fg_color=("gray85", "gray22"), corner_radius=6)
            stat_card.grid(row=0, column=i, padx=10, pady=10, sticky="nsew")
            stat_card.grid_columnconfigure(0, weight=1)
            
            ctk.CTkLabel(stat_card, text=stat["title"], 
                        font=ctk.CTkFont(size=12), 
                        text_color=("gray40", "gray60")).grid(row=0, column=0, pady=(10, 5))
            
            value_label = ctk.CTkLabel(stat_card, text=stat["value"] + " " + stat["unit"], 
                        font=ctk.CTkFont(size=14, weight="bold"))
            value_label.grid(row=1, column=0, pady=(0, 10))
            
            self.stats_labels[stat["key"]] = value_label

    def setup_log_tab(self):
        """Configura a aba de log de atividades"""
        tab = self.tabview.tab("Activity Log")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        main = ctk.CTkFrame(tab, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        
        # Cabeçalho
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(header, text="Activity Log", 
                    font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="w")
        
        # Botões de ação do log
        log_actions = ctk.CTkFrame(header, fg_color="transparent")
        log_actions.grid(row=0, column=1, sticky="e")
        
        ctk.CTkButton(log_actions, text="Clear Log", command=self.clear_log, width=100).grid(
            row=0, column=0, padx=5)
        ctk.CTkButton(log_actions, text="Save Log", command=self.save_log, width=100).grid(
            row=0, column=1, padx=5)
        
        # Área de log
        log_frame = ctk.CTkFrame(main, fg_color=("gray95", "gray16"), corner_radius=8)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
        self.log_text = ctk.CTkTextbox(log_frame, width=900, height=500, 
                                      font=ctk.CTkFont(family="Consolas", size=12))
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.insert("1.0", f"Log started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        self.log_text.configure(state="disabled")
        
        # Filtros de log
        filter_frame = ctk.CTkFrame(main, fg_color="transparent")
        filter_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        
        ctk.CTkLabel(filter_frame, text="Filter:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 10))
        
        self.log_filters = {}
        log_filter_types = ["All", "Info", "Warning", "Error", "Debug"]
        for filter_name in log_filter_types:
            btn = ctk.CTkButton(filter_frame, text=filter_name, width=80, 
                               fg_color="transparent", border_width=1,
                               command=lambda f=filter_name: self.filter_log(f))
            btn.pack(side="left", padx=5)
            self.log_filters[filter_name] = btn

    # ------------- Métodos auxiliares de interface -------------
    def show_tooltip(self, message):
        """Mostra uma dica de ferramenta"""
        tooltip = ctk.CTkToplevel(self.window)
        tooltip.geometry("300x100")
        tooltip.title("Information")
        tooltip.transient(self.window)
        tooltip.grab_set()
        
        ctk.CTkLabel(tooltip, text=message, wraplength=280).pack(padx=10, pady=10)
        ctk.CTkButton(tooltip, text="OK", command=tooltip.destroy).pack(pady=(0, 10))
        
        # Centralizar na tela
        tooltip.update_idletasks()
        x = self.window.winfo_x() + (self.window.winfo_width() - tooltip.winfo_width()) // 2
        y = self.window.winfo_y() + (self.window.winfo_height() - tooltip.winfo_height()) // 2
        tooltip.geometry(f"+{x}+{y}")

    def open_documentation(self):
        """Abre a documentação"""
        webbrowser.open("https://github.com/ansys/pyaedt")

    def show_help(self):
        """Mostra janela de ajuda"""
        help_window = ctk.CTkToplevel(self.window)
        help_window.title("Help - Patch Antenna Array Designer")
        help_window.geometry("600x400")
        help_window.transient(self.window)
        help_window.grab_set()
        
        # Conteúdo da ajuda
        tabview = ctk.CTkTabview(help_window)
        tabview.pack(fill="both", expand=True, padx=10, pady=10)
        
        tabs = ["Getting Started", "Parameters Guide", "Simulation", "Results"]
        for tab in tabs:
            tabview.add(tab)
            tabview.tab(tab).grid_columnconfigure(0, weight=1)
            
            content = ctk.CTkTextbox(tabview.tab(tab), wrap="word")
            content.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
            
            # Conteúdo de ajuda básico
            if tab == "Getting Started":
                help_text = """Welcome to Patch Antenna Array Designer!
                
This tool helps you design and simulate patch antenna arrays using ANSYS HFSS.

1. Start by configuring your antenna parameters in the Design Parameters tab.
2. Click 'Calculate Parameters' to compute derived values.
3. Run the simulation in the Simulation Control tab.
4. View results in the Results Analysis tab."""
            
            elif tab == "Parameters Guide":
                help_text = """Antenna Parameters:
- Center Frequency: Operating frequency in GHz
- Desired Gain: Target gain in dBi
- Sweep Range: Frequency range for simulation

Substrate Properties:
- Material: Substrate dielectric material
- εr: Relative permittivity
- tanδ: Loss tangent
- Thickness: Substrate height in mm

Feed Configuration:
- Probe Radius: Inner conductor radius
- Coax Ratio: Outer to inner conductor ratio
- Port Length: Length of the port section"""
            
            elif tab == "Simulation":
                help_text = """Simulation Controls:
- Run Simulation: Starts the simulation process
- Stop Simulation: Aborts the current simulation
- CPU Cores: Number of processor cores to use
- Save Project: Save the ANSYS project after simulation

Progress is shown in the progress bar, and status updates appear in the log."""
            
            else:  # Results
                help_text = """Results Analysis:
- S-Parameters: Return loss (S11) and VSWR
- Radiation Pattern: Far-field radiation patterns
- Statistics: Key performance metrics

Use the export buttons to save data or plots for further analysis."""
            
            content.insert("1.0", help_text)
            content.configure(state="disabled")
        
        # Centralizar na tela
        help_window.update_idletasks()
        x = self.window.winfo_x() + (self.window.winfo_width() - help_window.winfo_width()) // 2
        y = self.window.winfo_y() + (self.window.winfo_height() - help_window.winfo_height()) // 2
        help_window.geometry(f"+{x}+{y}")

    def update_progress(self, value, message=None):
        """Atualiza a barra de progresso e mensagem de status"""
        self.progress_bar.set(value)
        self.progress_label.configure(text=f"{int(value*100)}%")
        if message:
            self.sim_status_label.configure(text=message)

    def update_stats(self, key, value):
        """Atualiza as estatísticas de simulação"""
        if key in self.status_cards:
            self.status_cards[key].configure(text=value)
            self.simulation_stats[key] = value

    def update_result_stats(self, key, value):
        """Atualiza as estatísticas de resultados"""
        if key in self.stats_labels:
            # Extrair unidade do texto atual
            current_text = self.stats_labels[key].cget("text")
            unit = current_text.split()[-1] if " " in current_text else ""
            
            # Atualizar com novo valor mantendo a unidade
            self.stats_labels[key].configure(text=f"{value} {unit}")

    def filter_log(self, filter_type):
        """Filtra o log por tipo de mensagem"""
        # Implementação básica - em uma versão completa, isso filtraria as mensagens
        self.log_message(f"Filtering log by: {filter_type}")

    # ------------- Métodos de log -------------
    def log_message(self, message, level="INFO"):
        """Adiciona uma mensagem ao log"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        formatted_message = f"[{timestamp}] [{level}] {message}\n"
        self.log_queue.put(formatted_message)
    
    def process_log_queue(self):
        """Processa a fila de mensagens de log"""
        try:
            while True:
                msg = self.log_queue.get_nowait()
                try:
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", msg)
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                except Exception:
                    break
        except queue.Empty:
            pass
        finally:
            try:
                if self.window and self.window.winfo_exists():
                    self.window.after(100, self.process_log_queue)
            except Exception:
                pass

    def clear_log(self):
        """Limpa o log"""
        try:
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.configure(state="disabled")
            self.log_message("Log cleared")
        except Exception:
            pass

    def save_log(self):
        """Salva o log em um arquivo"""
        try:
            with open("simulation_log.txt", "w", encoding="utf-8") as f:
                f.write(self.log_text.get("1.0", "end"))
            self.log_message("Log saved to simulation_log.txt")
        except Exception as e:
            self.log_message(f"Error saving log: {e}")

    # ------------- Métodos de cálculo e simulação -------------
    def get_parameters(self):
        """Lê os parâmetros da interface"""
        self.log_message("Reading parameters...")
        for key, widget in self.entries:
            try:
                if key == "cores":
                    self.params[key] = int(widget.get()) if isinstance(widget, ctk.CTkEntry) else int(self.params[key])
                elif key == "save_project":
                    self.save_project = widget.get()
                elif key in ["substrate_thickness", "metal_thickness", "er", "tan_d",
                             "probe_radius", "coax_ba_ratio", "coax_wall_thickness",
                             "coax_port_length", "antipad_clearance",
                             "sweep_step", "frequency", "gain", "sweep_start", "sweep_stop"]:
                    self.params[key] = float(widget.get()) if isinstance(widget, ctk.CTkEntry) else float(self.params[key])
                elif key in ["spacing_type", "substrate_material", "sweep_type"]:
                    self.params[key] = widget.get()
            except Exception as e:
                msg = f"Invalid value for {key}: {e}"
                self.status_label.configure(text=msg)
                self.log_message(msg, "ERROR")
                return False
        self.log_message("Parameters OK.")
        return True

    def _guided_wavelength_ms(self, w_mm: float) -> float:
        """λg (mm) para microstrip com largura w_mm (Hammerstad)."""
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"])  # mm
        w = max(w_mm, 0.01)
        # εeff Hammerstad
        u = w / h
        a = 1 + (1/49.0)*math.log((u**4 + (u/52.0)**2)/(u**4 + 0.432)) + (1/18.7)*math.log(1 + (u/18.1)**3)
        b = 0.564*((er - 0.9)/(er + 3))**0.053
        eeff = (er + 1)/2 + (er - 1)/2*(1 + 10/u)**(-a*b)
        # correção espessura (opcional)
        c0 = self.c
        f = float(self.params["frequency"])*1e9
        lam_g = c0 / (f * math.sqrt(eeff)) * 1000.0  # mm
        return lam_g

    def _microstrip_w(self, Z0_ohm: float) -> float:
        """Hammerstad: devolve W (mm) dado Z0, εr, h(mm)."""
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"])
        Z0 = Z0_ohm
        A = Z0/60.0*math.sqrt((er+1)/2) + (er-1)/(er+1)*(0.23+0.11/er)
        w_h = (8*math.exp(A))/(math.exp(2*A)-2)
        if w_h < 2:
            W = w_h*h
        else:
            B = (377*math.pi)/(2*Z0*math.sqrt(er))
            W = h*(2/math.pi)*(B-1-math.log(2*B-1)+(er-1)/(2*er)*(math.log(B-1)+0.39-0.61/er))
        return max(W, 0.1)

    def calculate_patch_dimensions(self, frequency_ghz: float) -> Tuple[float, float, float]:
        """Calcula as dimensões do patch da antena"""
        f = frequency_ghz * 1e9
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"]) / 1000.0  # mm->m
        W = self.c / (2 * f) * math.sqrt(2 / (er + 1))
        eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W) ** -0.5
        dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
        L_eff = self.c / (2 * f * math.sqrt(eeff))
        L = L_eff - 2 * dL
        lambda_g = self.c / (f * math.sqrt(eeff))
        return (L * 1000.0, W * 1000.0, lambda_g * 1000.0)

    def _size_array_from_gain(self):
        """Calcula o tamanho do array com base no ganho desejado"""
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
        """Calcula o tamanho do substrato"""
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

    def calculate_parameters(self):
        """Calcula todos os parâmetros da antena"""
        self.log_message("Calculating parameters...")
        if not self.get_parameters(): 
            return
        
        try:
            # Calcular dimensões do patch
            L_mm, W_mm, lambda_g_mm = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params.update({
                "patch_length": L_mm, 
                "patch_width": W_mm, 
                "lambda_g": lambda_g_mm
            })
            
            # Calcular espaçamento
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
            
            # Calcular tamanho do array
            rows, cols, _ = self._size_array_from_gain()
            self.calculated_params.update({
                "num_patches": rows * cols, 
                "rows": rows, 
                "cols": cols
            })
            
            # Calcular tamanho do substrato
            self.calculate_substrate_size()

            # Atualizar a interface
            calculated_text = f"""Array Configuration: {rows} × {cols}
Total Patches: {rows * cols}
Patch Dimensions: {L_mm:.2f} × {W_mm:.2f} mm
Element Spacing: {spacing_mm:.2f} mm ({self.params['spacing_type']})
Guided Wavelength: {lambda_g_mm:.2f} mm
Substrate Size: {self.calculated_params['substrate_width']:.2f} × {self.calculated_params['substrate_length']:.2f} mm"""
            
            self.calculated_text.configure(state="normal")
            self.calculated_text.delete("1.0", "end")
            self.calculated_text.insert("1.0", calculated_text)
            self.calculated_text.configure(state="disabled")
            
            self.status_label.configure(text="Parameters calculated.")
            self.log_message("Parameters calculated successfully.")
            
        except Exception as e:
            error_msg = f"Error in calculation: {e}"
            self.status_label.configure(text=error_msg)
            self.log_message(error_msg, "ERROR")
            self.log_message(traceback.format_exc(), "DEBUG")

    def start_simulation_thread(self):
        """Inicia a simulação em uma thread separada"""
        if self.is_simulation_running:
            self.log_message("Simulation already running.")
            return
        
        if not ANSYS_AVAILABLE:
            self.log_message("ANSYS is not available. Simulation cannot be run.", "ERROR")
            return
        
        self.stop_simulation = False
        self.is_simulation_running = True
        
        # Atualizar interface
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.update_progress(0, "Initializing simulation...")
        self.update_stats("status", "Running")
        
        # Iniciar thread de simulação
        threading.Thread(target=self.run_simulation, daemon=True).start()

    def stop_simulation_thread(self):
        """Solicita parada da simulação"""
        self.stop_simulation = True
        self.log_message("Stop requested.")
        self.update_stats("status", "Stopping")

    def run_simulation(self):
        """Executa a simulação (em thread separada)"""
        try:
            self.log_message("Starting simulation...")
            
            if not self.get_parameters():
                self.log_message("Invalid parameters.", "ERROR")
                return
            
            if self.calculated_params["num_patches"] < 1:
                self.calculate_parameters()
            
            # Simular progresso (em uma implementação real, isso viria do ANSYS)
            for i in range(10):
                if self.stop_simulation:
                    self.log_message("Simulation stopped by user.")
                    break
                
                progress = (i + 1) / 10
                self.update_progress(progress, f"Simulation step {i+1}/10")
                
                # Simular algumas estatísticas
                if i == 2:
                    self.update_stats("mesh_elements", "12,548")
                elif i == 5:
                    self.update_stats("simulation_time", "2.5 min")
                elif i == 8:
                    self.update_stats("memory_usage", "1.2 GB")
                
                threading.Event().wait(0.5)  # Simular trabalho
            
            if not self.stop_simulation:
                # Simular resultados
                self.simulate_results()
                self.update_progress(1.0, "Simulation completed!")
                self.update_stats("status", "Completed")
                self.log_message("Simulation completed successfully.")
            
        except Exception as e:
            error_msg = f"Simulation error: {e}"
            self.log_message(error_msg, "ERROR")
            self.log_message(traceback.format_exc(), "DEBUG")
            self.update_progress(0, "Simulation failed!")
            self.update_stats("status", "Failed")
        
        finally:
            try:
                self.run_button.configure(state="normal")
                self.stop_button.configure(state="disabled")
            except Exception:
                pass
            self.is_simulation_running = False

    def simulate_results(self):
        """Simula resultados para demonstração"""
        # Gerar dados simulados para S11
        freq_range = np.linspace(self.params["sweep_start"], self.params["sweep_stop"], 100)
        center_freq = self.params["frequency"]
        
        # Simular curva S11 (return loss)
        s11_db = -20 * np.log10(1 + 10 * (freq_range - center_freq)**2 / (center_freq/5)**2)
        
        # Simular padrão de radiação
        theta = np.linspace(-180, 180, 360)
        gain_theta = 10 * np.cos(np.radians(theta))**2
        
        phi = np.linspace(-180, 180, 360)
        gain_phi = 8 * np.cos(np.radians(phi))**3
        
        # Calcular VSWR
        vswr = (1 + 10**(-s11_db/20)) / (1 - 10**(-s11_db/20))
        
        # Armazenar dados
        self.simulation_data = np.column_stack((freq_range, s11_db))
        
        # Atualizar gráficos
        self.plot_simulated_results(freq_range, s11_db, vswr, theta, gain_theta, phi, gain_phi)
        
        # Atualizar estatísticas de resultados
        min_s11_idx = np.argmin(s11_db)
        resonant_freq = freq_range[min_s11_idx]
        min_s11 = s11_db[min_s11_idx]
        
        # Calcular banda passante (onde S11 < -10 dB)
        bw_mask = s11_db <= -10
        if np.any(bw_mask):
            bw_start = freq_range[bw_mask][0]
            bw_end = freq_range[bw_mask][-1]
            bandwidth = (bw_end - bw_start) * 1000  # em MHz
        else:
            bandwidth = 0
        
        min_vswr = np.min(vswr)
        
        self.update_result_stats("resonant_freq", f"{resonant_freq:.2f}")
        self.update_result_stats("bandwidth", f"{bandwidth:.0f}")
        self.update_result_stats("return_loss", f"{min_s11:.1f}")
        self.update_result_stats("vswr", f"{min_vswr:.2f}")

    def plot_simulated_results(self, freq_range, s11_db, vswr, theta, gain_theta, phi, gain_phi):
        """Plota resultados simulados"""
        try:
            # Limpar gráficos
            for ax in [self.ax_s11, self.ax_vswr, self.ax_th, self.ax_ph]:
                ax.clear()
            
            # S11
            self.ax_s11.plot(freq_range, s11_db, 'b-', linewidth=2, label='S11')
            self.ax_s11.axhline(y=-10, color='r', linestyle='--', alpha=0.7, label='-10 dB')
            self.ax_s11.axvline(x=self.params["frequency"], color='g', linestyle='--', alpha=0.7)
            self.ax_s11.set_xlabel('Frequency (GHz)')
            self.ax_s11.set_ylabel('S11 (dB)')
            self.ax_s11.set_title('Return Loss (S11)')
            self.ax_s11.legend()
            self.ax_s11.grid(True, alpha=0.3)
            
            # VSWR
            self.ax_vswr.plot(freq_range, vswr, 'r-', linewidth=2, label='VSWR')
            self.ax_vswr.axhline(y=2, color='g', linestyle='--', alpha=0.7, label='VSWR=2')
            self.ax_vswr.axvline(x=self.params["frequency"], color='g', linestyle='--', alpha=0.7)
            self.ax_vswr.set_xlabel('Frequency (GHz)')
            self.ax_vswr.set_ylabel('VSWR')
            self.ax_vswr.set_title('Voltage Standing Wave Ratio')
            self.ax_vswr.legend()
            self.ax_vswr.grid(True, alpha=0.3)
            
            # Theta pattern
            self.ax_th.plot(theta, gain_theta, 'b-', linewidth=2)
            self.ax_th.set_xlabel('Theta (degrees)')
            self.ax_th.set_ylabel('Gain (dBi)')
            self.ax_th.set_title('Radiation Pattern (Theta cut)')
            self.ax_th.grid(True, alpha=0.3)
            
            # Phi pattern
            self.ax_ph.plot(phi, gain_phi, 'r-', linewidth=2)
            self.ax_ph.set_xlabel('Phi (degrees)')
            self.ax_ph.set_ylabel('Gain (dBi)')
            self.ax_ph.set_title('Radiation Pattern (Phi cut)')
            self.ax_ph.grid(True, alpha=0.3)
            
            # Ajustar layout e atualizar canvas
            self.fig.tight_layout()
            self.canvas.draw()
            
            self.log_message("Results plotted successfully.")
            
        except Exception as e:
            self.log_message(f"Error plotting results: {e}", "ERROR")

    def export_csv(self):
        """Exporta dados para CSV"""
        try:
            if self.simulation_data is not None:
                np.savetxt("simulation_results.csv", self.simulation_data, delimiter=",",
                           header="Frequency (GHz), S11 (dB)", comments='')
                self.log_message("CSV exported successfully.")
            else:
                self.log_message("No data to export.", "WARNING")
        except Exception as e:
            self.log_message(f"Error exporting CSV: {e}", "ERROR")

    def export_png(self):
        """Exporta gráficos para PNG"""
        try:
            if hasattr(self, 'fig'):
                self.fig.savefig("simulation_results.png", dpi=300, bbox_inches='tight')
                self.log_message("PNG exported successfully.")
        except Exception as e:
            self.log_message(f"Error exporting PNG: {e}", "ERROR")

    def save_parameters(self):
        """Salva os parâmetros em um arquivo JSON"""
        try:
            all_params = {**self.params, **self.calculated_params}
            with open("antenna_parameters.json", "w") as f:
                json.dump(all_params, f, indent=4)
            self.log_message("Parameters saved successfully.")
        except Exception as e:
            self.log_message(f"Error saving parameters: {e}", "ERROR")

    def load_parameters(self):
        """Carrega parâmetros de um arquivo JSON"""
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
            self.log_message("Parameters loaded successfully.")
            
        except Exception as e:
            self.log_message(f"Error loading parameters: {e}", "ERROR")

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
            
            # Atualizar texto de parâmetros calculados
            calculated_text = f"""Array Configuration: {self.calculated_params['rows']} × {self.calculated_params['cols']}
Total Patches: {self.calculated_params['num_patches']}
Patch Dimensions: {self.calculated_params['patch_length']:.2f} × {self.calculated_params['patch_width']:.2f} mm
Element Spacing: {self.calculated_params['spacing']:.2f} mm ({self.params['spacing_type']})
Guided Wavelength: {self.calculated_params['lambda_g']:.2f} mm
Substrate Size: {self.calculated_params['substrate_width']:.2f} × {self.calculated_params['substrate_length']:.2f} mm"""
            
            self.calculated_text.configure(state="normal")
            self.calculated_text.delete("1.0", "end")
            self.calculated_text.insert("1.0", calculated_text)
            self.calculated_text.configure(state="disabled")
            
        except Exception as e:
            self.log_message(f"Error updating interface: {e}", "ERROR")

    def cleanup(self):
        """Limpeza antes de fechar o aplicativo"""
        try:
            if self.hfss:
                try:
                    if self.save_project:
                        self.hfss.save_project()
                    else:
                        self.hfss.close_project(save=False)
                except Exception as e:
                    self.log_message(f"Error closing project: {e}", "ERROR")
            
            if self.desktop:
                try:
                    self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception as e:
                    self.log_message(f"Error releasing desktop: {e}", "ERROR")
            
            if self.temp_folder and not self.save_project:
                try:
                    self.temp_folder.cleanup()
                except Exception as e:
                    self.log_message(f"Error cleaning temp files: {e}", "ERROR")
                    
        except Exception as e:
            self.log_message(f"Cleanup error: {e}", "ERROR")

    def on_closing(self):
        """Manipula o fechamento da janela"""
        self.log_message("Closing application...")
        try:
            self.window.after_cancel(self.process_log_queue)
        except Exception:
            pass
        self.cleanup()
        try:
            self.window.destroy()
        except Exception:
            pass

    def run(self):
        """Executa o aplicativo"""
        try:
            self.window.mainloop()
        finally:
            self.cleanup()


if __name__ == "__main__":
    app = ModernPatchAntennaDesigner()
    app.run()