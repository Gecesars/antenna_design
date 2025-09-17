# -*- coding: utf-8 -*-
import os
import tempfile
from datetime import datetime
import math
import json
import traceback
import queue
import threading
from typing import Tuple, Optional, List

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk

from ansys.aedt.core import Desktop, Hfss


# ---------- Aparência ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class PatchArrayDesigner:
    def __init__(self):
        # AEDT
        self.desktop: Optional[Desktop] = None
        self.hfss: Optional[Hfss] = None
        self.temp_folder = None
        self.project_path = ""
        self.project_title = "patch_array"
        self.design_name = "patch_array"

        # GUI / Estado
        self.log_queue = queue.Queue()
        self.is_simulation_running = False
        self.save_project = False
        self.stop_simulation = False
        self.simulation_data = None
        self._shape_id = 0

        # ---------- Parâmetros ----------
        self.params = {
            "frequency": 10.0,             # GHz
            "gain": 12.0,                  # dBi (alvo)
            "sweep_start": 8.0,
            "sweep_stop": 12.0,
            "sweep_type": "Interpolating",
            "sweep_step": 0.02,            # GHz
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": False,

            # Array (preenchidos após calcular/confirmar)
            "rows": 2,
            "cols": 2,
            "spacing_type": "lambda/2",

            # Substrato
            "substrate_material": "Duroid (tm)",
            "er": 2.2,
            "tan_d": 0.0009,
            "substrate_thickness": 0.5,    # mm
            "metal_thickness": 0.035,      # mm

            # Coax
            "probe_radius": 0.40,          # mm (a)
            "coax_ba_ratio": 2.3,          # b/a
            "coax_wall_thickness": 0.20,   # mm
            "coax_port_length": 3.0,       # mm
            "antipad_clearance": 0.10,     # mm

            # Conexões
            "overlap": 0.10,               # mm que entra no patch
            "neck_len": 0.8,               # mm pescoço dentro do patch
        }

        # ---------- Calculados ----------
        self.calc = {
            "num_patches": 4,
            "patch_length": 9.5,           # mm
            "patch_width": 9.3,            # mm
            "spacing": 6.0,                # mm (entre bordas)
            "substrate_width": 40.0,
            "substrate_length": 40.0,
            "lambda_g50": 0.0,
            "W50": 0.0, "W70": 0.0, "W100": 0.0, "W200": 0.0
        }

        self.c = 299792458.0
        self._build_gui()

    # ========================= GUI =========================
    def _build_gui(self):
        self.win = ctk.CTk()
        self.win.title("Patch Antenna Array Designer")
        self.win.geometry("1300x900")
        self.win.grid_columnconfigure(0, weight=1)
        self.win.grid_rowconfigure(1, weight=1)

        # Header
        head = ctk.CTkFrame(self.win, height=64, corner_radius=16)
        head.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        head.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            head, text="Patch Antenna Array Designer",
            font=ctk.CTkFont(size=26, weight="bold")
        ).grid(row=0, column=0, pady=12)

        # Tabs
        self.tabs = ctk.CTkTabview(self.win, segmented_button_selected_color="#2E8B57")
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        for t in ["Parâmetros", "Simulação", "Resultados", "Log"]:
            self.tabs.add(t)
            self.tabs.tab(t).grid_columnconfigure(0, weight=1)

        self._build_params_tab()
        self._build_sim_tab()
        self._build_results_tab()
        self._build_log_tab()

        # Status
        status = ctk.CTkFrame(self.win, height=40, corner_radius=12)
        status.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 10))
        status.grid_propagate(False)
        self.status = ctk.CTkLabel(status, text="Pronto.", font=ctk.CTkFont(weight="bold"))
        self.status.pack(pady=6)

        # Log loop
        self._process_log_async()
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

    def _sec(self, parent, title):
        f = ctk.CTkFrame(parent, corner_radius=12)
        ctk.CTkLabel(f, text=title, font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=12, pady=(10, 6)
        )
        ctk.CTkFrame(f, height=2).pack(fill="x", padx=10, pady=(0, 8))
        return f

    def _build_params_tab(self):
        tab = self.tabs.tab("Parâmetros")
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure((0, 1), weight=1)

        left = ctk.CTkScrollableFrame(tab)
        right = ctk.CTkScrollableFrame(tab)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=6)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=6)

        self.entries: List[tuple] = []

        def add_entry(frame, label, key, value, combo=None, check=False, width=160):
            row = ctk.CTkFrame(frame)
            row.pack(fill="x", padx=10, pady=4)
            ctk.CTkLabel(row, text=label).pack(side="left")
            if combo:
                var = ctk.StringVar(value=str(value))
                w = ctk.CTkComboBox(row, values=[str(v) for v in combo], variable=var, width=width)
                w.pack(side="right"); self.entries.append((key, var))
            elif check:
                var = ctk.BooleanVar(value=value)
                w = ctk.CTkCheckBox(row, text="", variable=var)
                w.pack(side="right"); self.entries.append((key, var))
            else:
                w = ctk.CTkEntry(row, width=width); w.insert(0, str(value)); w.pack(side="right")
                self.entries.append((key, w))

        # Antena / Array
        sec = self._sec(left, "Antena / Array")
        sec.pack(fill="x", padx=6, pady=6)
        add_entry(sec, "F0 (GHz):", "frequency", self.params["frequency"])
        add_entry(sec, "Ganho desejado (dBi):", "gain", self.params["gain"])
        add_entry(sec, "Sweep início (GHz):", "sweep_start", self.params["sweep_start"])
        add_entry(sec, "Sweep fim (GHz):", "sweep_stop", self.params["sweep_stop"])
        add_entry(sec, "Sweep:", "sweep_type", self.params["sweep_type"], combo=["Discrete", "Interpolating", "Fast"])
        add_entry(sec, "Passo Discrete (GHz):", "sweep_step", self.params["sweep_step"])
        # rows/cols agora são decididos pela caixa de diálogo – mas mantemos aqui para inspeção/ajuste rápido
        add_entry(sec, "Linhas (rows):", "rows", self.params["rows"], combo=["1", "2"])
        add_entry(sec, "Colunas (cols):", "cols", self.params["cols"], combo=["1", "2"])
        add_entry(sec, "Espaçamento:", "spacing_type", self.params["spacing_type"],
                  combo=["lambda/2", "0.7*lambda", "0.8*lambda", "0.9*lambda", "lambda"])

        # Substrato
        sec = self._sec(left, "Substrato")
        sec.pack(fill="x", padx=6, pady=6)
        add_entry(sec, "Material:", "substrate_material", self.params["substrate_material"],
                  combo=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"])
        add_entry(sec, "εr:", "er", self.params["er"])
        add_entry(sec, "tanδ:", "tan_d", self.params["tan_d"])
        add_entry(sec, "Espessura (mm):", "substrate_thickness", self.params["substrate_thickness"])
        add_entry(sec, "Metal (mm):", "metal_thickness", self.params["metal_thickness"])

        # Coax
        sec = self._sec(left, "Coax")
        sec.pack(fill="x", padx=6, pady=6)
        add_entry(sec, "Raio interno a (mm):", "probe_radius", self.params["probe_radius"])
        add_entry(sec, "Razão b/a:", "coax_ba_ratio", self.params["coax_ba_ratio"])
        add_entry(sec, "Parede blindagem (mm):", "coax_wall_thickness", self.params["coax_wall_thickness"])
        add_entry(sec, "Comprimento porto (mm):", "coax_port_length", self.params["coax_port_length"])
        add_entry(sec, "Clear antipad (mm):", "antipad_clearance", self.params["antipad_clearance"])

        # Conexões
        sec = self._sec(left, "Conexões no patch")
        sec.pack(fill="x", padx=6, pady=6)
        add_entry(sec, "Overlap no patch (mm):", "overlap", self.params["overlap"])
        add_entry(sec, "Neck (mm):", "neck_len", self.params["neck_len"])

        # Simulação
        sec = self._sec(left, "Simulação")
        sec.pack(fill="x", padx=6, pady=6)
        add_entry(sec, "CPU cores:", "cores", self.params["cores"])
        add_entry(sec, "Salvar projeto:", "save_project", self.save_project, check=True)

        # Calculados
        sec = self._sec(right, "Calculados / Dimensões")
        sec.pack(fill="x", padx=6, pady=6)
        self.lbl_np = ctk.CTkLabel(sec, text="Patches: -- (defina o ganho e clique Calcular)")
        self.lbl_np.pack(anchor="w", padx=10, pady=4)
        self.lbl_dims = ctk.CTkLabel(sec, text="Patch: -- x -- mm")
        self.lbl_dims.pack(anchor="w", padx=10, pady=4)
        self.lbl_spacing = ctk.CTkLabel(sec, text="Espaçamento: -- mm")
        self.lbl_spacing.pack(anchor="w", padx=10, pady=4)
        self.lbl_sub = ctk.CTkLabel(sec, text="Substrato: -- x -- mm")
        self.lbl_sub.pack(anchor="w", padx=10, pady=4)
        self.lbl_ws = ctk.CTkLabel(sec, text="W50/W70/W100/W200: -- / -- / -- / -- mm")
        self.lbl_ws.pack(anchor="w", padx=10, pady=4)

        # Botões
        sec = ctk.CTkFrame(right, corner_radius=12)
        sec.pack(fill="x", padx=6, pady=6)
        ctk.CTkButton(sec, text="Calcular",
                      command=self.calculate_parameters, fg_color="#2E8B57", height=38).pack(
            side="left", padx=8, pady=8)
        ctk.CTkButton(sec, text="Salvar parâmetros",
                      command=self.save_parameters, fg_color="#4169E1", height=38).pack(
            side="left", padx=8, pady=8)
        ctk.CTkButton(sec, text="Carregar",
                      command=self.load_parameters, fg_color="#FF8C00", height=38).pack(
            side="left", padx=8, pady=8)

    def _build_sim_tab(self):
        tab = self.tabs.tab("Simulação")
        tab.grid_columnconfigure(0, weight=1)

        ctl = ctk.CTkFrame(tab, corner_radius=12)
        ctl.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        ctk.CTkLabel(ctl, text="Controle de simulação",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)

        row = ctk.CTkFrame(ctl); row.pack(pady=8)
        self.btn_run = ctk.CTkButton(row, text="Executar", fg_color="#2E8B57",
                                     command=self._start_sim, width=160, height=40)
        self.btn_run.pack(side="left", padx=6)
        self.btn_stop = ctk.CTkButton(row, text="Parar", fg_color="#DC143C",
                                      command=self._stop_sim, state="disabled",
                                      width=160, height=40)
        self.btn_stop.pack(side="left", padx=6)

        pf = ctk.CTkFrame(ctl); pf.pack(fill="x", padx=40, pady=8)
        ctk.CTkLabel(pf, text="Progresso:").pack(anchor="w")
        self.progress = ctk.CTkProgressBar(pf, height=18); self.progress.pack(fill="x", pady=6); self.progress.set(0)
        self.lbl_state = ctk.CTkLabel(ctl, text="Aguardando…", font=ctk.CTkFont(weight="bold"))
        self.lbl_state.pack(pady=6)

    def _build_results_tab(self):
        tab = self.tabs.tab("Resultados")
        tab.grid_rowconfigure(1, weight=1); tab.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(tab, text="Resultados",
                     font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=8)

        fr = ctk.CTkFrame(tab, corner_radius=12)
        fr.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        fr.grid_rowconfigure(0, weight=1); fr.grid_columnconfigure(0, weight=1)

        self.fig = plt.figure(figsize=(10, 9))
        face = '#2B2B2B' if ctk.get_appearance_mode() == "Dark" else '#FFFFFF'
        self.fig.patch.set_facecolor(face)
        self.ax_s11 = self.fig.add_subplot(3, 1, 1)
        self.ax_th = self.fig.add_subplot(3, 1, 2)
        self.ax_ph = self.fig.add_subplot(3, 1, 3)
        for ax in [self.ax_s11, self.ax_th, self.ax_ph]:
            ax.set_facecolor(face); ax.grid(True, alpha=0.4)
        self.canvas = FigureCanvasTkAgg(self.fig, master=fr)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        ex = ctk.CTkFrame(tab, corner_radius=12)
        ex.grid(row=2, column=0, pady=6)
        ctk.CTkButton(ex, text="Exportar CSV", command=self._export_csv, fg_color="#6A5ACD").pack(side="left", padx=6)
        ctk.CTkButton(ex, text="Exportar PNG", command=self._export_png, fg_color="#20B2AA").pack(side="left", padx=6)

    def _build_log_tab(self):
        tab = self.tabs.tab("Log")
        tab.grid_rowconfigure(1, weight=1); tab.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(tab, text="Log",
                     font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=8)
        self.log_box = ctk.CTkTextbox(tab, width=900, height=520, font=ctk.CTkFont(family="Consolas"))
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_box.insert("1.0", "Log iniciado em " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        btns = ctk.CTkFrame(tab); btns.grid(row=2, column=0, pady=8)
        ctk.CTkButton(btns, text="Limpar", command=lambda: self.log_box.delete("1.0", "end")).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Salvar", command=self._save_log).pack(side="left", padx=6)

    # ===================== Util/Log =====================
    def _log(self, msg: str):
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")

    def _process_log_async(self):
        try:
            while True:
                m = self.log_queue.get_nowait()
                try:
                    self.log_box.insert("end", m); self.log_box.see("end")
                except Exception:
                    break
        except queue.Empty:
            pass
        finally:
            try:
                if self.win and self.win.winfo_exists():
                    self.win.after(100, self._process_log_async)
            except Exception:
                pass

    # ===================== Cálculos RF =====================
    def _get_params(self) -> bool:
        try:
            for key, w in self.entries:
                if key in ["rows", "cols"]:
                    self.params[key] = int(w.get()) if isinstance(w, ctk.StringVar) else int(w.get())
                elif key == "save_project":
                    self.save_project = w.get()
                elif key in ["frequency", "gain", "sweep_start", "sweep_stop", "sweep_step",
                             "er", "tan_d", "substrate_thickness", "metal_thickness",
                             "probe_radius", "coax_ba_ratio", "coax_wall_thickness",
                             "coax_port_length", "antipad_clearance", "overlap", "neck_len"]:
                    self.params[key] = float(w.get()) if isinstance(w, ctk.CTkEntry) else float(self.params[key])
                elif key in ["substrate_material", "sweep_type", "spacing_type"]:
                    self.params[key] = w.get()
                elif key == "cores":
                    self.params[key] = int(w.get()) if isinstance(w, ctk.CTkEntry) else int(self.params[key])
            return True
        except Exception as e:
            self.status.configure(text=f"Valor inválido: {e}")
            self._log(f"Valor inválido: {e}")
            return False

    def _microstrip_w(self, Z0: float) -> float:
        er = float(self.params["er"]); h = float(self.params["substrate_thickness"])
        A = Z0/60.0*math.sqrt((er+1)/2.0) + (er-1)/(er+1)*(0.23+0.11/er)
        wh = (8*math.exp(A))/(math.exp(2*A)-2)
        if wh < 2:
            W = wh*h
        else:
            B = (377*math.pi)/(2*Z0*math.sqrt(er))
            W = h*(2/math.pi)*(B - 1 - math.log(2*B-1) + (er-1)/(2*er)*(math.log(B-1)+0.39-0.61/er))
        return max(W, 0.08)

    def _guided_lambda(self, w_mm: float) -> float:
        er = float(self.params["er"])
        h = float(self.params["substrate_thickness"])
        u = max(w_mm, 0.01) / h
        a = 1 + (1/49.0)*math.log((u**4 + (u/52.0)**2)/(u**4 + 0.432)) + (1/18.7)*math.log(1 + (u/18.1)**3)
        b = 0.564*((er - 0.9)/(er + 3))**0.053
        eeff = (er + 1)/2 + (er - 1)/2*(1 + 10/u)**(-a*b)
        f = float(self.params["frequency"])*1e9
        return self.c/(f*math.sqrt(eeff))*1000.0

    def _patch_dims(self, f_ghz: float) -> Tuple[float, float, float]:
        f = f_ghz*1e9; er = float(self.params["er"])
        h = float(self.params["substrate_thickness"])/1000.0
        W = self.c/(2*f)*math.sqrt(2/(er+1.0))
        eeff = (er+1)/2 + (er-1)/2*(1 + 12*h/W)**-0.5
        dL = 0.412*h*((eeff+0.3)*(W/h+0.264))/((eeff-0.258)*(W/h+0.8))
        L_eff = self.c/(2*f*math.sqrt(eeff))
        L = L_eff - 2*dL
        lamg = self.c/(f*math.sqrt(eeff))
        return L*1000.0, W*1000.0, lamg*1000.0

    def _pairs_for_N(self, N: int) -> List[Tuple[int, int]]:
        pairs = []
        for r in range(1, int(math.sqrt(N)) + 1):
            if N % r == 0:
                pairs.append((r, N // r))
        # adicionar espelhados somente se diferentes
        pairs = list(dict.fromkeys(pairs + [(c, r) for r, c in pairs if r != c]))
        return sorted(pairs)

    def _pick_layout_dialog(self, N: int) -> Tuple[int, int]:
        """Abre um diálogo para o usuário escolher m×n dado N. Retorna (rows, cols)."""
        supported = {(1, 2), (2, 1), (2, 2)}   # redes implementadas
        options = self._pairs_for_N(N)
        if not options:
            options = [(1, N)]

        dlg = ctk.CTkToplevel(self.win)
        dlg.title("Escolher configuração m × n")
        dlg.geometry("420x250"); dlg.resizable(False, False)
        ctk.CTkLabel(dlg, text=f"Número de patches calculado: {N}",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(14, 6))
        msg = ("Escolha a configuração m×n.\n"
               "Itens marcados com ★ são suportados pela rede atual (1×2, 2×1, 2×2).")
        ctk.CTkLabel(dlg, text=msg, wraplength=380, justify="center").pack(pady=4)

        nice = []
        for r, c in options:
            tag = " ★" if (r, c) in supported else ""
            nice.append(f"{r} x {c}{tag}")

        var = ctk.StringVar(value=nice[0])
        cb = ctk.CTkComboBox(dlg, values=nice, variable=var, width=180)
        cb.pack(pady=10)

        chosen = {"rows": options[0][0], "cols": options[0][1]}

        def confirm():
            s = var.get().split("x")
            r = int(s[0].strip()); c = int(s[1].split("★")[0].strip())
            chosen["rows"], chosen["cols"] = r, c
            dlg.destroy()

        ctk.CTkButton(dlg, text="Confirmar", command=confirm, fg_color="#2E8B57").pack(pady=10)
        dlg.grab_set(); dlg.wait_window()
        return chosen["rows"], chosen["cols"]

    def _compute_array(self, ask_layout: bool = True):
        f0 = float(self.params["frequency"])
        G_des = float(self.params["gain"])
        G_elem = 8.0  # ~ patch alimentado por sonda
        N_req = max(1, int(math.ceil(10 ** ((G_des - G_elem) / 10.0))))

        # Por simetria e rede corporativa simples, favorecemos números pares
        if N_req == 1:
            N_req = 2
        if N_req % 2 == 1:
            N_req += 1

        # Pergunta configuração m×n ao usuário
        if ask_layout:
            rows, cols = self._pick_layout_dialog(N_req)
            self.params["rows"], self.params["cols"] = rows, cols
        r = int(self.params["rows"]); c = int(self.params["cols"])
        self.calc["num_patches"] = r * c

        L, W, lamg50 = self._patch_dims(f0)
        self.calc["patch_length"] = L
        self.calc["patch_width"] = W
        self.calc["lambda_g50"] = lamg50

        # spacing (entre bordas)
        lam0_mm = self.c/(f0*1e9)*1000.0
        k = {"lambda/2": 0.5, "0.7*lambda": 0.7, "0.8*lambda": 0.8, "0.9*lambda": 0.9, "lambda": 1.0}[self.params["spacing_type"]]
        s = k*lam0_mm
        self.calc["spacing"] = s

        # Substrato
        tot_w = c*W + (c-1)*s
        tot_l = r*L + (r-1)*s
        margin = max(tot_w, tot_l)*0.20
        self.calc["substrate_width"] = tot_w + 2*margin
        self.calc["substrate_length"] = tot_l + 2*margin

        # Larguras
        self.calc["W50"]  = self._microstrip_w(50.0)
        self.calc["W70"]  = self._microstrip_w(70.710678)
        self.calc["W100"] = self._microstrip_w(100.0)
        self.calc["W200"] = self._microstrip_w(200.0)

    def calculate_parameters(self):
        self._log("Calculando parâmetros…")
        if not self._get_params():
            return
        try:
            self._compute_array(ask_layout=True)
            self.lbl_np.configure(text=f"Patches: {self.calc['num_patches']}  (config: {self.params['rows']} × {self.params['cols']})")
            self.lbl_dims.configure(text=f"Patch: {self.calc['patch_length']:.2f} × {self.calc['patch_width']:.2f} mm")
            self.lbl_spacing.configure(text=f"Espaçamento (borda-borda): {self.calc['spacing']:.2f} mm ({self.params['spacing_type']})")
            self.lbl_sub.configure(text=f"Substrato: {self.calc['substrate_width']:.1f} × {self.calc['substrate_length']:.1f} mm")
            self.lbl_ws.configure(text=f"W50/W70/W100/W200: {self.calc['W50']:.3f} / {self.calc['W70']:.3f} / {self.calc['W100']:.3f} / {self.calc['W200']:.3f} mm")
            self.status.configure(text="Parâmetros calculados.")
        except Exception as e:
            self._log(f"Erro no cálculo: {e}\n{traceback.format_exc()}")
            self.status.configure(text=f"Erro: {e}")

    # ===================== AEDT Helpers =====================
    def _ensure_material(self, name: str, er: float, tan_d: float):
        try:
            if not self.hfss.materials.checkifmaterialexists(name):
                self.hfss.materials.add_material(name)
                m = self.hfss.materials.material_keys[name]
                m.permittivity = er
                m.dielectric_loss_tangent = tan_d
                self._log(f"Material criado: {name}")
        except Exception as e:
            self._log(f"Material warn '{name}': {e}")

    def _safe_name(self, base: str) -> str:
        self._shape_id += 1
        return f"{base}_{self._shape_id:04d}"

    def _rect_h(self, x1, x2, y, w, name):
        x_min, x_max = min(x1, x2), max(x1, x2)
        if abs(x_max - x_min) < 1e-6:
            x_max = x_min + 1e-6
        return self.hfss.modeler.create_rectangle("XY", [x_min, y - w/2, "h_sub"],
                                                  [x_max - x_min, w],
                                                  name=self._safe_name(name), material="copper")

    def _rect_v(self, y1, y2, x, w, name):
        y_min, y_max = min(y1, y2), max(y1, y2)
        if abs(y_max - y_min) < 1e-6:
            y_max = y_min + 1e-6
        return self.hfss.modeler.create_rectangle("XY", [x - w/2, y_min, "h_sub"],
                                                  [w, y_max - y_min],
                                                  name=self._safe_name(name), material="copper")

    def _create_coax_lumped(self, ground, substrate, x0: float, y0: float, tag: str):
        """Pino interno atravessa; blindagem para no GND; porto anelar no -Lp."""
        try:
            a = float(self.params["probe_radius"])
            b = max(a*float(self.params["coax_ba_ratio"]), a + 0.02)
            wall = float(self.params["coax_wall_thickness"])
            Lp = float(self.params["coax_port_length"])
            h = float(self.params["substrate_thickness"])
            clr = float(self.params["antipad_clearance"])

            pin = self.hfss.modeler.create_cylinder("Z", [x0, y0, -Lp], a, h + Lp + 1e-3,
                                                    name=self._safe_name(f"{tag}_Pin"), material="copper")
            shield = self.hfss.modeler.create_cylinder("Z", [x0, y0, -Lp], b + wall, Lp,
                                                       name=self._safe_name(f"{tag}_Shield"), material="copper")
            void = self.hfss.modeler.create_cylinder("Z", [x0, y0, -Lp], b, Lp,
                                                     name=self._safe_name(f"{tag}_Void"), material="vacuum")
            self.hfss.modeler.subtract(shield, [void], keep_originals=False)

            # Antipad e furo no GND
            hole_r = b + clr
            subhole = self.hfss.modeler.create_cylinder("Z", [x0, y0, 0], hole_r, h,
                                                        name=self._safe_name(f"{tag}_SubHole"), material="vacuum")
            self.hfss.modeler.subtract(substrate, [subhole], keep_originals=False)
            gndhole = self.hfss.modeler.create_circle("XY", [x0, y0, 0], hole_r,
                                                      name=self._safe_name(f"{tag}_GHole"), material="vacuum")
            self.hfss.modeler.subtract(ground, [gndhole], keep_originals=False)

            # Porto anelar
            ring = self.hfss.modeler.create_circle("XY", [x0, y0, -Lp], b,
                                                   name=self._safe_name(f"{tag}_Ring"), material="vacuum")
            inner = self.hfss.modeler.create_circle("XY", [x0, y0, -Lp], a,
                                                    name=self._safe_name(f"{tag}_Hole"), material="vacuum")
            self.hfss.modeler.subtract(ring, [inner], keep_originals=False)

            # linha de integração radial
            r1 = a + 0.02; r2 = b - 0.02
            p1 = [x0 + r1, y0, -Lp]; p2 = [x0 + r2, y0, -Lp]
            _ = self.hfss.lumped_port(assignment=ring.name, integration_line=[p1, p2],
                                      impedance=50.0, name=f"{tag}_Lumped", renormalize=True)

            # pad superior (solda na trilha)
            top_pad = self.hfss.modeler.create_circle("XY", [x0, y0, "h_sub"], a,
                                                      name=self._safe_name(f"{tag}_TopPad"), material="copper")
            self._log(f"Lumped Port {tag}_Lumped criado.")
            return top_pad
        except Exception as e:
            self._log(f"Coax erro: {e}\n{traceback.format_exc()}")
            return None

    # ===================== Simulação =====================
    def _start_sim(self):
        if self.is_simulation_running:
            self._log("Simulação já em execução.")
            return
        self.stop_simulation = False
        self.is_simulation_running = True
        threading.Thread(target=self._run_sim, daemon=True).start()

    def _stop_sim(self):
        self.stop_simulation = True
        self._log("Parada solicitada.")

    def _open_project(self):
        self._log("Abrindo projeto…")
        if self.desktop is None:
            self.desktop = Desktop(
                version=self.params["aedt_version"],
                non_graphical=self.params["non_graphical"],
                new_desktop=True
            )
            self._log("Desktop inicializado.")

        if self.temp_folder is None:
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.project_path = os.path.join(self.temp_folder.name, f"{self.project_title}_{ts}.aedt")

        oDesktop = getattr(self.desktop, "odesktop", None) or getattr(self.desktop, "_odesktop", None)
        oDesktop.NewProject()
        oProject = oDesktop.GetActiveProject()
        if self.design_name not in [d.GetName() for d in oProject.GetDesigns()]:
            oProject.InsertDesign("HFSS", self.design_name, "DrivenModal", "")
        oProject.SetActiveDesign(self.design_name)
        try: oProject.SaveAs(self.project_path, True)
        except Exception: pass

        self.hfss = Hfss(
            project=oProject.GetName(),
            design=self.design_name,
            solution_type="DrivenModal",
            version=self.params["aedt_version"],
            non_graphical=self.params["non_graphical"],
            new_desktop=False
        )
        self._log(f"Projeto ativo: {self.project_path}")

    def _run_sim(self):
        try:
            self._log("Iniciando simulação…")
            self.btn_run.configure(state="disabled"); self.btn_stop.configure(state="normal")
            self.lbl_state.configure(text="Rodando…"); self.progress.set(0)

            if not self._get_params():
                self._log("Parâmetros inválidos."); return
            # ao iniciar sim direto, não incomodar o usuário — usar rows/cols já escolhidos
            self._compute_array(ask_layout=False)

            self._open_project()
            self.progress.set(0.2)
            self.hfss.modeler.model_units = "mm"

            # Materiais
            sub_name = self.params["substrate_material"]
            if not self.hfss.materials.checkifmaterialexists(sub_name):
                sub_name = "Custom_Substrate"
                self._ensure_material(sub_name, float(self.params["er"]), float(self.params["tan_d"]))

            # Variáveis
            L = self.calc["patch_length"]; W = self.calc["patch_width"]; s = self.calc["spacing"]
            r = int(self.params["rows"]); c = int(self.params["cols"])
            subW = self.calc["substrate_width"]; subL = self.calc["substrate_length"]
            h = float(self.params["substrate_thickness"])

            self.hfss["h_sub"] = f"{h}mm"
            self.hfss["patchL"] = f"{L}mm"; self.hfss["patchW"] = f"{W}mm"; self.hfss["spacing"] = f"{s}mm"
            self.hfss["subW"] = f"{subW}mm"; self.hfss["subL"] = f"{subL}mm"
            self.hfss["OLAP"] = f"{self.params['overlap']}mm"

            # Substrato e Ground
            substrate = self.hfss.modeler.create_box(["-subW/2", "-subL/2", 0], ["subW", "subL", "h_sub"],
                                                     name=self._safe_name("Substrate"), material=sub_name)
            ground = self.hfss.modeler.create_rectangle("XY", ["-subW/2", "-subL/2", 0], ["subW", "subL"],
                                                        name=self._safe_name("Ground"), material="copper")

            # Posições de patches (centradas)
            xs = [ (k - (c-1)/2)*(W + s) for k in range(c) ]
            ys = [ (i - (r-1)/2)*(L + s) for i in range(r) ]

            patches = []
            for cy in ys:
                for cx in xs:
                    p = self.hfss.modeler.create_rectangle(
                        "XY", [cx - W/2, cy - L/2, "h_sub"], ["patchW", "patchL"],
                        name=self._safe_name("Patch"), material="copper")
                    patches.append({"cx": cx, "cy": cy, "obj": p})

            # Porta coax + pad (no centro)
            top_pad = self._create_coax_lumped(ground, substrate, 0.0, 0.0, "P0")

            # Larguras e λ/4
            W50 = self.calc["W50"]; W70 = self.calc["W70"]; W100 = self.calc["W100"]; W200 = self.calc["W200"]
            Lq70 = self._guided_lambda(W70)/4.0
            self.hfss["W50"] = f"{W50}mm"; self.hfss["W70"] = f"{W70}mm"
            self.hfss["W100"] = f"{W100}mm"; self.hfss["W200"] = f"{W200}mm"; self.hfss["Lq70"] = f"{Lq70}mm"

            copper_objs = [top_pad] if top_pad else []

            # ======= Rede corporativa (suporta 1x2, 2x1, 2x2) =======
            overlap = float(self.params["overlap"])
            neck = float(self.params["neck_len"])

            def inner_edge_x(cx):
                return (cx - W/2 + overlap) if cx > 0 else (cx + W/2 - overlap)

            def inner_edge_y(cy):
                return (cy - L/2 + overlap) if cy > 0 else (cy + L/2 - overlap)

            supported = {(1, 2), (2, 1), (2, 2)}
            if (r, c) not in supported:
                self._log(f"Aviso: rede corporativa atual suporta 1x2, 2x1, 2x2. Usando melhor aproximação para {r}x{c}.")
            # 2x2
            if r == 2 and c == 2:
                xL, xR = xs[0], xs[1]
                # Tronco horizontal 70 Ω λ/4 + braços 100 Ω
                copper_objs += [
                    self._rect_h(0.0, -Lq70, 0.0, W70, "Q70_L"),
                    self._rect_h(0.0, +Lq70, 0.0, W70, "Q70_R"),
                ]
                LhL = abs(xL) - Lq70; LhR = abs(xR) - Lq70
                LhL = LhL if LhL > 0 else 0.2
                LhR = LhR if LhR > 0 else 0.2
                copper_objs += [
                    self._rect_h(-Lq70, -Lq70 - LhL, 0.0, W100, "ARM_L"),
                    self._rect_h(+Lq70, +Lq70 + LhR, 0.0, W100, "ARM_R"),
                    self._rect_v(0.0, +Lq70, xL, W70, "Q70U_L"),
                    self._rect_v(0.0, -Lq70, xL, W70, "Q70D_L"),
                    self._rect_v(0.0, +Lq70, xR, W70, "Q70U_R"),
                    self._rect_v(0.0, -Lq70, xR, W70, "Q70D_R"),
                ]
                yT, yB = ys[1], ys[0]
                yT_edge = inner_edge_y(yT); yB_edge = inner_edge_y(yB)
                copper_objs += [
                    self._rect_v(+Lq70, yT_edge, xL, W200, "ST200U_L"),
                    self._rect_v(+Lq70, yT_edge, xR, W200, "ST200U_R"),
                    self._rect_v(yT_edge, yT_edge + neck, xL, W200, "NECKU_L"),
                    self._rect_v(yT_edge, yT_edge + neck, xR, W200, "NECKU_R"),
                    self._rect_v(-Lq70, yB_edge, xL, W200, "ST200D_L"),
                    self._rect_v(-Lq70, yB_edge, xR, W200, "ST200D_R"),
                    self._rect_v(yB_edge - neck, yB_edge, xL, W200, "NECKD_L"),
                    self._rect_v(yB_edge - neck, yB_edge, xR, W200, "NECKD_R"),
                ]

            # 1x2
            elif r == 1 and c == 2:
                xL, xR = xs[0], xs[1]; y0 = ys[0]
                copper_objs += [
                    self._rect_h(0.0, -Lq70, y0, W70, "Q70_L"),
                    self._rect_h(0.0, +Lq70, y0, W70, "Q70_R"),
                ]
                LhL = abs(xL) - Lq70; LhR = abs(xR) - Lq70
                LhL = LhL if LhL > 0 else 0.2
                LhR = LhR if LhR > 0 else 0.2
                copper_objs += [
                    self._rect_h(-Lq70, -Lq70 - LhL, y0, W100, "ARM_L"),
                    self._rect_h(+Lq70, +Lq70 + LhR, y0, W100, "ARM_R"),
                ]
                xL_edge = inner_edge_x(xL); xR_edge = inner_edge_x(xR)
                copper_objs += [
                    self._rect_h(xL, xL_edge, y0, W200, "ST200_L"),
                    self._rect_h(xL_edge - neck, xL_edge, y0, W200, "NECK_L"),
                    self._rect_h(xR, xR_edge, y0, W200, "ST200_R"),
                    self._rect_h(xR_edge, xR_edge + neck, y0, W200, "NECK_R"),
                ]

            # 2x1
            elif r == 2 and c == 1:
                x0 = xs[0]
                copper_objs += [
                    self._rect_v(0.0, +Lq70, 0.0, W70, "Q70_UP"),
                    self._rect_v(0.0, -Lq70, 0.0, W70, "Q70_DN"),
                ]
                yT, yB = ys[1], ys[0]
                LvU = abs(yT) - Lq70; LvD = abs(yB) - Lq70
                LvU = LvU if LvU > 0 else 0.2
                LvD = LvD if LvD > 0 else 0.2
                copper_objs += [
                    self._rect_v(+Lq70, +Lq70 + LvU, 0.0, W100, "ARM_UP"),
                    self._rect_v(-Lq70, -Lq70 - LvD, 0.0, W100, "ARM_DN"),
                ]
                yT_edge = inner_edge_y(yT); yB_edge = inner_edge_y(yB)
                copper_objs += [
                    self._rect_v(+Lq70 + LvU, yT_edge, 0.0, W200, "ST200_UP"),
                    self._rect_v(yT_edge, yT_edge + neck, 0.0, W200, "NECK_UP"),
                    self._rect_v(-Lq70 - LvD, yB_edge, 0.0, W200, "ST200_DN"),
                    self._rect_v(yB_edge - neck, yB_edge, 0.0, W200, "NECK_DN"),
                ]

            # Unir topo
            all_top = [p["obj"] for p in patches] + [o for o in copper_objs if o]
            try:
                uni = self.hfss.modeler.unite(all_top)
                top_name = uni.name if hasattr(uni, "name") else None
            except Exception as e:
                top_name = None
                self._log(f"Unite aviso: {e}")

            self.progress.set(0.55)

            # Boundaries
            try:
                names = [ground.name]
                if top_name: names.append(top_name)
                else: names += [o.name for o in all_top]
                self.hfss.assign_perfecte_to_sheets(list(dict.fromkeys(names)))
                self._log(f"PerfectE em: {names}")
            except Exception as e:
                self._log(f"PerfectE aviso: {e}")

            # Região + radiação
            self._log("Criando região + radiação")
            lam0_mm = self.c/(float(self.params["sweep_start"])*1e9)*1000.0
            pad = max(5.0, min(0.25*lam0_mm, 15.0))
            region = self.hfss.modeler.create_region([pad]*6, is_percentage=False)
            self.hfss.assign_radiation_boundary_to_objects(region)

            # Infinite Sphere (se possível)
            try:
                rf = self.hfss.odesign.GetModule("RadField")
                rf.InsertInfiniteSphereSetup([
                    "NAME:IS1",
                    "UseCustomRadiationSurface:=", False,
                    "CSDefinition:=", "Theta-Phi",
                    "Polarization:=", "Linear",
                    "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "1deg",
                    "PhiStart:=", "-180deg", "PhiStop:=", "180deg", "PhiStep:=", "1deg",
                    "UseLocalCS:=", False
                ])
            except Exception as e:
                self._log(f"InfiniteSphere aviso: {e}")

            self.progress.set(0.65)

            # Setup + Sweep
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.params['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02
            try:
                stype = self.params["sweep_type"]
                if stype == "Discrete":
                    setup.create_linear_step_sweep(unit="GHz", start_frequency=self.params["sweep_start"],
                                                   stop_frequency=self.params["sweep_stop"],
                                                   step_size=float(self.params["sweep_step"]), name="Sweep1")
                elif stype == "Fast":
                    setup.create_frequency_sweep(unit="GHz", name="Sweep1",
                                                 start_frequency=self.params["sweep_start"],
                                                 stop_frequency=self.params["sweep_stop"], sweep_type="Fast")
                else:
                    setup.create_frequency_sweep(unit="GHz", name="Sweep1",
                                                 start_frequency=self.params["sweep_start"],
                                                 stop_frequency=self.params["sweep_stop"], sweep_type="Interpolating")
            except Exception as e:
                self._log(f"Sweep aviso: {e}")

            if self.save_project: self.hfss.save_project()
            self.hfss.analyze_setup("Setup1", cores=self.params["cores"])

            # Pós-processamento
            self.progress.set(0.9)
            self._plot_results()
            self.progress.set(1.0)
            self.lbl_state.configure(text="Concluída")
            self._log("Concluído.")
        except Exception as e:
            self._log(f"Erro geral: {e}\n{traceback.format_exc()}")
            self.lbl_state.configure(text=f"Erro: {e}")
        finally:
            try:
                self.btn_run.configure(state="normal")
                self.btn_stop.configure(state="disabled")
            except Exception:
                pass
            self.is_simulation_running = False

    # ===================== Resultados =====================
    def _plot_results(self):
        try:
            self._log("Plotando…")
            for ax in [self.ax_s11, self.ax_th, self.ax_ph]:
                ax.clear(); ax.grid(True, alpha=0.4)

            # S11
            expr = "dB(S(1,1))"
            sol = None
            try:
                rpt = self.hfss.post.reports_by_category.standard(expressions=[expr])
                rpt.context = ["Setup1: Sweep1"]
                sol = rpt.get_solution_data()
            except Exception as e:
                self._log(f"S11 aviso: {e}")

            if sol:
                try:
                    f = np.asarray(sol.primary_sweep_values, dtype=float)
                    y = np.asarray(sol.data_real()[0], dtype=float)
                    if f.size and f.size == y.size:
                        self.simulation_data = np.column_stack((f, y))
                        self.ax_s11.plot(f, y, linewidth=2, label="S11")
                        self.ax_s11.axhline(-10, linestyle="--", alpha=0.6, label="-10 dB")
                        self.ax_s11.axvline(float(self.params["frequency"]), linestyle="--", alpha=0.6)
                        self.ax_s11.set_xlabel("Freq (GHz)"); self.ax_s11.set_ylabel("dB"); self.ax_s11.set_title("S11"); self.ax_s11.legend()
                    else:
                        self.ax_s11.text(0.5, 0.5, "S11 indisponível", transform=self.ax_s11.transAxes,
                                         ha="center", va="center")
                except Exception as e:
                    self._log(f"S11 parse erro: {e}")
                    self.ax_s11.text(0.5, 0.5, "S11 indisponível", transform=self.ax_s11.transAxes,
                                     ha="center", va="center")
            else:
                self.ax_s11.text(0.5, 0.5, "S11 indisponível", transform=self.ax_s11.transAxes,
                                 ha="center", va="center")

            # Far-Field cortes (se existir)
            try:
                def cut(kind, fixed):
                    rn = f"Gain_{kind}_{fixed}"
                    ctx = "IS1"
                    if kind == "theta":
                        prim = "Theta"; vars_ = {"Freq": f"{self.params['frequency']}GHz", "Phi": f"{fixed}deg", "Theta": "All"}
                    else:
                        prim = "Phi";   vars_ = {"Freq": f"{self.params['frequency']}GHz", "Theta": f"{fixed}deg", "Phi": "All"}
                    rep = self.hfss.post.reports_by_category.far_field(
                        expressions=["dB(GainTotal)"], context=ctx,
                        primary_sweep_variable=prim,
                        setup="Setup1 : LastAdaptive", variations=vars_, name=rn
                    )
                    sd = rep.get_solution_data()
                    return np.array(sd.primary_sweep_values, dtype=float), np.array(sd.data_real())[0]

                th, gth = cut("theta", 0.0)
                self.ax_th.plot(th, gth); self.ax_th.set_xlabel("Theta (°)"); self.ax_th.set_ylabel("dB"); self.ax_th.set_title("FF: Theta @ Phi=0")
                ph, gph = cut("phi", 90.0)
                self.ax_ph.plot(ph, gph); self.ax_ph.set_xlabel("Phi (°)"); self.ax_ph.set_ylabel("dB"); self.ax_ph.set_title("FF: Phi @ Theta=90")
            except Exception as e:
                self._log(f"Far-field indisponível: {e}")
                self.ax_th.text(0.5, 0.5, "FF indisponível", transform=self.ax_th.transAxes, ha="center", va="center")
                self.ax_ph.text(0.5, 0.5, "FF indisponível", transform=self.ax_ph.transAxes, ha="center", va="center")

            self.fig.tight_layout()
            self.canvas.draw()
            self._log("Plot OK.")
        except Exception as e:
            self._log(f"Erro nos gráficos: {e}\n{traceback.format_exc()}")

    # ===================== Persistência =====================
    def save_parameters(self):
        try:
            data = {**self.params, **self.calc}
            with open("antenna_parameters.json", "w") as f:
                json.dump(data, f, indent=2)
            self._log("Parâmetros salvos.")
        except Exception as e:
            self._log(f"Salvar parâmetros erro: {e}")

    def load_parameters(self):
        try:
            with open("antenna_parameters.json", "r") as f:
                data = json.load(f)
            for k in self.params:
                if k in data:
                    self.params[k] = data[k]
            # Recarrega na GUI
            for k, w in self.entries:
                if k in self.params:
                    if isinstance(w, ctk.CTkEntry):
                        w.delete(0, "end"); w.insert(0, str(self.params[k]))
                    elif isinstance(w, ctk.StringVar):
                        w.set(str(self.params[k]))
                    elif isinstance(w, ctk.BooleanVar):
                        w.set(bool(self.params[k]))
            # Recalcula e pergunta layout de novo (útil após carregar)
            self._compute_array(ask_layout=True)
            self.lbl_np.configure(text=f"Patches: {self.calc['num_patches']}  (config: {self.params['rows']} × {self.params['cols']})")
            self.lbl_dims.configure(text=f"Patch: {self.calc['patch_length']:.2f} × {self.calc['patch_width']:.2f} mm")
            self.lbl_spacing.configure(text=f"Espaçamento (borda-borda): {self.calc['spacing']:.2f} mm ({self.params['spacing_type']})")
            self.lbl_sub.configure(text=f"Substrato: {self.calc['substrate_width']:.1f} × {self.calc['substrate_length']:.1f} mm")
            self.lbl_ws.configure(text=f"W50/W70/W100/W200: {self.calc['W50']:.3f} / {self.calc['W70']:.3f} / {self.calc['W100']:.3f} / {self.calc['W200']:.3f} mm")
            self._log("Parâmetros carregados.")
        except Exception as e:
            self._log(f"Carregar parâmetros erro: {e}")

    # ===================== Export/Util =====================
    def _export_csv(self):
        try:
            if self.simulation_data is not None:
                np.savetxt("simulation_results.csv", self.simulation_data, delimiter=",",
                           header="Frequency (GHz), S11 (dB)", comments="")
                self._log("CSV exportado.")
            else:
                self._log("Sem dados para exportar.")
        except Exception as e:
            self._log(f"CSV erro: {e}")

    def _export_png(self):
        try:
            self.fig.savefig("simulation_results.png", dpi=300, bbox_inches="tight")
            self._log("PNG salvo.")
        except Exception as e:
            self._log(f"PNG erro: {e}")

    def _save_log(self):
        try:
            with open("simulation_log.txt", "w", encoding="utf-8") as f:
                f.write(self.log_box.get("1.0", "end"))
            self._log("Log salvo.")
        except Exception as e:
            self._log(f"Salvar log erro: {e}")

    # ===================== Encerramento =====================
    def _cleanup(self):
        try:
            if self.hfss:
                try:
                    if self.save_project: self.hfss.save_project()
                    else: self.hfss.close_project(save=False)
                except Exception as e: self._log(f"Close projeto erro: {e}")
            if self.desktop:
                try: self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception as e: self._log(f"Desktop release erro: {e}")
            if self.temp_folder and not self.save_project:
                try: self.temp_folder.cleanup()
                except Exception as e: self._log(f"Temp cleanup erro: {e}")
        except Exception as e:
            self._log(f"Cleanup erro: {e}")

    def _on_close(self):
        self._log("Fechando…")
        try:
            self.win.after_cancel(self._process_log_async)
        except Exception:
            pass
        self._cleanup()
        try:
            self.win.destroy()
        except Exception:
            pass

    def run(self):
        try:
            self.win.mainloop()
        finally:
            self._cleanup()


if __name__ == "__main__":
    app = PatchArrayDesigner()
    app.run()
