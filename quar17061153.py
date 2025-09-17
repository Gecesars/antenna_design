# -*- coding: utf-8 -*-
"""
Array de patches paramétrico (1x1 a 10x10) com rede corporativa em fase,
GUI moderna (CustomTkinter) e automação PyAEDT.

Requisitos:
  pip install customtkinter matplotlib numpy
  pip install ansys-aedt-core==0.19.0
  HFSS 2024 R2+ instalado

Autor: você + ChatGPT (2025)
"""

import os, math, json, tempfile, traceback, threading, queue, re
from datetime import datetime
from typing import List, Tuple, Optional, Dict

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk

from ansys.aedt.core import Desktop, Hfss

# ---------- Aparência ----------
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# ---------- EM utils ----------
C0 = 299_792_458.0

def hammerstad_w(er: float, h_mm: float, z0: float) -> float:
    """Largura (mm) de microstrip para Z0, εr, h(mm)."""
    h = float(h_mm)
    A = z0 / 60.0 * math.sqrt((er + 1) / 2) + (er - 1) / (er + 1) * (0.23 + 0.11 / er)
    w_h = (8 * math.exp(A)) / (math.exp(2 * A) - 2)
    if w_h < 2:
        W = w_h * h
    else:
        B = (377 * math.pi) / (2 * z0 * math.sqrt(er))
        W = h * (2 / math.pi) * (B - 1 - math.log(2 * B - 1)
             + (er - 1) / (2 * er) * (math.log(B - 1) + 0.39 - 0.61 / er))
    return max(W, 0.05)

def eff_eps(er: float, h_mm: float, w_mm: float) -> float:
    """εef Hammerstad."""
    h = float(h_mm)
    w = max(w_mm, 1e-3)
    u = w / h
    a = 1 + (1 / 49.0) * math.log((u**4 + (u/52.0)**2)/(u**4 + 0.432)) + (1/18.7) * math.log(1 + (u/18.1)**3)
    b = 0.564 * ((er - 0.9)/(er + 3))**0.053
    return (er + 1) / 2 + (er - 1) / 2 * (1 + 10 / u)**(-a * b)

def guided_lambda_mm(freq_ghz: float, er: float, h_mm: float, w_mm: float) -> float:
    ee = eff_eps(er, h_mm, w_mm)
    return C0 / (freq_ghz * 1e9 * math.sqrt(ee)) * 1000.0

def patch_dims_er(f_ghz: float, er: float, h_mm: float) -> Tuple[float, float, float]:
    """Dimensões de patch (L, W) e λg aproximado (mm)."""
    f = f_ghz * 1e9
    h = h_mm / 1000.0
    W = C0 / (2 * f) * math.sqrt(2 / (er + 1))
    eeff = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * h / W)**-0.5
    dL = 0.412 * h * ((eeff + 0.3) * (W / h + 0.264)) / ((eeff - 0.258) * (W / h + 0.8))
    L_eff = C0 / (2 * f * math.sqrt(eeff))
    L = L_eff - 2 * dL
    lamg = C0 / (f * math.sqrt(eeff))
    return L * 1000.0, W * 1000.0, lamg * 1000.0

def suggest_rows_cols_from_gain(g_des_dbi: float, max_side: int = 10) -> Tuple[int, int, int]:
    """Sugere N≈10^((Gdes-8)/10) e fatora em m×n."""
    g_elem = 8.0
    n_req = max(1, int(round(10 ** ((g_des_dbi - g_elem) / 10.0))))
    n_req = min(n_req, max_side * max_side)
    root = int(round(math.sqrt(n_req)))
    r = max(1, min(max_side, root))
    c = max(1, min(max_side, int(math.ceil(n_req / r))))
    while r * c < n_req and (r < max_side or c < max_side):
        if r <= c and r < max_side: r += 1
        elif c < max_side: c += 1
        else: break
    return r, c, r * c

# ---------- App ----------
class PatchArrayApp:
    def __init__(self):
        # AEDT
        self.desktop: Optional[Desktop] = None
        self.hfss: Optional[Hfss] = None
        self.tempdir: Optional[tempfile.TemporaryDirectory] = None

        # Estado
        self.rows = 2
        self.cols = 2
        self.n_total = 4
        self._id = 0
        self._portname = "P0_Lumped"

        # Parâmetros do usuário (GUI)
        self.p = {
            "frequency": 10.0,
            "gain": 12.0,
            "sweep_start": 8.0,
            "sweep_stop": 12.0,
            "sweep_type": "Interpolating",
            "sweep_step": 0.02,

            "substrate_material": "Duroid (tm)",  # se não existir, cria Custom_Substrate
            "er": 2.2,
            "tan_d": 9e-4,
            "h_sub": 0.5,     # mm
            "t_cu": 0.035,    # mm

            "spacing_scale": 0.5,  # múltiplo de λ0 (X=Y)
            "margin_factor": 0.2,  # margem do substrato

            # alimentação
            "probe_radius": 0.40,   # mm
            "coax_ba_ratio": 2.3,
            "coax_wall": 0.20,      # mm
            "coax_Lp": 3.0,         # mm
            "antipad": 0.10,        # mm

            # pequenas sobreposições para garantir conexão de metal
            "ovl": 0.08,            # mm
        }

        # Calculados / variáveis do design
        self.calc: Dict[str, float] = {}

        # GUI infra
        self.log_q = queue.Queue()
        self.sim_running = False
        self.sim_data = None

        self._build_gui()

    # ---------- GUI ----------
    def _build_gui(self):
        self.win = ctk.CTk()
        self.win.title("Patch Array Designer — Paramétrico")
        self.win.geometry("1380x900")
        self.win.grid_columnconfigure(0, weight=1)
        self.win.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.win, height=64)
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=10)
        ctk.CTkLabel(header, text="Patch Array Designer • HFSS",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(side="left", padx=10, pady=12)

        self.tabs = ctk.CTkTabview(self.win)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        for t in ["Parâmetros", "Simulação", "Resultados", "Log"]:
            self.tabs.add(t)
            self.tabs.tab(t).grid_columnconfigure(0, weight=1)

        self._tab_params()
        self._tab_sim()
        self._tab_results()
        self._tab_log()

        status = ctk.CTkFrame(self.win, height=36)
        status.grid(row=2, column=0, sticky="ew", padx=12, pady=(0,8))
        status.grid_propagate(False)
        self.lb_status = ctk.CTkLabel(status, text="Pronto.", font=ctk.CTkFont(weight="bold"))
        self.lb_status.pack(pady=6)

        self._log_pump()
        self.win.protocol("WM_DELETE_WINDOW", self._close)

    def _section(self, parent, title, row, col=0):
        fr = ctk.CTkFrame(parent)
        fr.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)
        fr.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(fr, text=title, font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 6))
        ctk.CTkFrame(fr, height=2, fg_color=("#a0a0a0", "#2a2a2a")).grid(
            row=1, column=0, sticky="ew", padx=6, pady=(0,6))
        return fr

    def _entry(self, parent, label, key, row, width=140):
        ctk.CTkLabel(parent, text=label).grid(row=row, column=0, padx=12, pady=6, sticky="w")
        ent = ctk.CTkEntry(parent, width=width)
        ent.insert(0, str(self.p[key] if key in self.p else ""))
        ent.grid(row=row, column=1, padx=12, pady=6, sticky="w")
        setattr(self, f"ent_{key}", ent)
        return ent

    def _tab_params(self):
        tab = self.tabs.tab("Parâmetros")
        tab.grid_rowconfigure(0, weight=1)
        main = ctk.CTkScrollableFrame(tab)
        main.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        # Antena
        sec = self._section(main, "Antena", 0)
        r = 2
        self._entry(sec, "Frequência (GHz)", "frequency", r); r+=1
        self._entry(sec, "Ganho desejado (dBi)", "gain", r); r+=1
        self._entry(sec, "Sweep início (GHz)", "sweep_start", r); r+=1
        self._entry(sec, "Sweep fim (GHz)", "sweep_stop", r); r+=1

        row = ctk.CTkFrame(sec); row.grid(row=r, column=0, sticky="w", padx=10, pady=4); r+=1
        ctk.CTkLabel(row, text="Espaçamento (λ0 ×):").pack(side="left")
        self.cb_spacing = ctk.CTkComboBox(row, values=["0.5", "0.7", "0.8", "0.9", "1.0"], width=100)
        self.cb_spacing.set(str(self.p["spacing_scale"])); self.cb_spacing.pack(side="left", padx=8)

        # Layout
        secL = self._section(main, "Layout do Array", 1)
        rr = 2
        self.lb_sug = ctk.CTkLabel(secL, text="Sugestão: --", font=ctk.CTkFont(weight="bold")); self.lb_sug.grid(row=rr, column=0, padx=12, pady=(6,2), sticky="w"); rr+=1

        rc1 = ctk.CTkFrame(secL); rc1.grid(row=rr, column=0, padx=10, pady=6, sticky="w"); rr+=1
        ctk.CTkLabel(rc1, text="Linhas (m): ").pack(side="left")
        self.sl_rows = ctk.CTkSlider(rc1, from_=1, to=10, number_of_steps=9, width=220)
        self.sl_rows.set(self.rows); self.sl_rows.pack(side="left", padx=8)
        self.lb_rows = ctk.CTkLabel(rc1, text=str(self.rows)); self.lb_rows.pack(side="left")
        self.sl_rows.bind("<ButtonRelease-1>", lambda e: self.lb_rows.configure(text=f"{int(round(self.sl_rows.get()))}"))

        rc2 = ctk.CTkFrame(secL); rc2.grid(row=rr, column=0, padx=10, pady=6, sticky="w"); rr+=1
        ctk.CTkLabel(rc2, text="Colunas (n): ").pack(side="left")
        self.sl_cols = ctk.CTkSlider(rc2, from_=1, to=10, number_of_steps=9, width=220)
        self.sl_cols.set(self.cols); self.sl_cols.pack(side="left", padx=8)
        self.lb_cols = ctk.CTkLabel(rc2, text=str(self.cols)); self.lb_cols.pack(side="left")
        self.sl_cols.bind("<ButtonRelease-1>", lambda e: self.lb_cols.configure(text=f"{int(round(self.sl_cols.get()))}"))

        bt = ctk.CTkFrame(secL); bt.grid(row=rr, column=0, sticky="w", padx=10, pady=(6,8)); rr+=1
        ctk.CTkButton(bt, text="Sugerir m×n pelo ganho", command=self._suggest_mn).pack(side="left", padx=6)
        ctk.CTkButton(bt, text="Aplicar", command=self._apply_mn, fg_color="#22c55e").pack(side="left", padx=6)

        self.lb_dims = ctk.CTkLabel(secL, text="Patch: -- × -- mm | λg(50 Ω): -- mm | Substrato: -- × -- mm")
        self.lb_dims.grid(row=rr, column=0, padx=12, pady=(8,6), sticky="w")

        # Substrato
        secS = self._section(main, "Substrato e Metal", 2)
        rs = 2
        ctk.CTkLabel(secS, text="Material").grid(row=rs, column=0, padx=12, pady=6, sticky="w")
        self.cb_mat = ctk.CTkComboBox(secS, values=["Duroid (tm)", "Rogers RO4003C (tm)", "FR4_epoxy", "Air"], width=220)
        self.cb_mat.set(self.p["substrate_material"]); self.cb_mat.grid(row=rs, column=1, padx=12, pady=6, sticky="w"); rs+=1
        self._entry(secS, "εr (se custom)", "er", rs); rs+=1
        self._entry(secS, "tanδ", "tan_d", rs); rs+=1
        self._entry(secS, "Espessura h (mm)", "h_sub", rs); rs+=1
        self._entry(secS, "Metal t (mm)", "t_cu", rs); rs+=1

        # Alimentação
        secC = self._section(main, "Alimentação (coax + stub/linhas)", 3); rc=2
        self._entry(secC, "Raio interno a (mm)", "probe_radius", rc); rc+=1
        self._entry(secC, "Razão b/a", "coax_ba_ratio", rc); rc+=1
        self._entry(secC, "Parede shield (mm)", "coax_wall", rc); rc+=1
        self._entry(secC, "Comprimento porto Lp (mm)", "coax_Lp", rc); rc+=1
        self._entry(secC, "Antipad clearance (mm)", "antipad", rc); rc+=1
        self._entry(secC, "Sobreposição OVL (mm)", "ovl", rc); rc+=1

        # Simulação
        secP = self._section(main, "Simulação", 4); rp=2
        self._entry(secP, "Cores CPU", "cores", rp); rp+=1
        ctk.CTkLabel(secP, text="Sweep").grid(row=rp, column=0, padx=12, pady=6, sticky="w")
        self.cb_sweep = ctk.CTkComboBox(secP, values=["Interpolating", "Discrete", "Fast"], width=160)
        self.cb_sweep.set(self.p["sweep_type"]); self.cb_sweep.grid(row=rp, column=1, padx=12, pady=6, sticky="w"); rp+=1
        self._entry(secP, "Passo Discrete (GHz)", "sweep_step", rp); rp+=1

        self.ch_save = ctk.CTkCheckBox(secP, text="Salvar projeto AEDT", onvalue=True, offvalue=False)
        self.ch_save.grid(row=rp, column=0, padx=12, pady=6, sticky="w")

        # Ações
        act = ctk.CTkFrame(main); act.grid(row=5, column=0, sticky="w", padx=8, pady=10)
        ctk.CTkButton(act, text="Calcular", command=self._calculate, fg_color="#2563eb").pack(side="left", padx=6)
        ctk.CTkButton(act, text="Salvar parâmetros", command=self._save_params).pack(side="left", padx=6)
        ctk.CTkButton(act, text="Carregar parâmetros", command=self._load_params).pack(side="left", padx=6)

    def _tab_sim(self):
        tab = self.tabs.tab("Simulação")
        top = ctk.CTkFrame(tab); top.grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.bt_run = ctk.CTkButton(top, text="Executar", command=self._start, fg_color="#16a34a", width=160, height=40)
        self.bt_run.pack(side="left", padx=6)
        self.bt_stop = ctk.CTkButton(top, text="Parar", command=self._stop, state="disabled", fg_color="#dc2626", width=160, height=40)
        self.bt_stop.pack(side="left", padx=6)

        pr = ctk.CTkFrame(tab); pr.grid(row=1, column=0, sticky="ew", padx=10, pady=6)
        ctk.CTkLabel(pr, text="Progresso:").pack(anchor="w")
        self.pb = ctk.CTkProgressBar(pr, height=16); self.pb.pack(fill="x", padx=4, pady=4); self.pb.set(0)
        self.lb_sim = ctk.CTkLabel(tab, text="Aguardando…", font=ctk.CTkFont(weight="bold")); self.lb_sim.grid(row=2, column=0, padx=10, pady=6, sticky="w")

    def _tab_results(self):
        tab = self.tabs.tab("Resultados")
        tab.grid_rowconfigure(1, weight=1)
        ct = ctk.CTkFrame(tab); ct.grid(row=0, column=0, sticky="ew", padx=10, pady=(10,0))
        ctk.CTkLabel(ct, text="Resultados", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w")

        plot = ctk.CTkFrame(tab); plot.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        plot.grid_columnconfigure(0, weight=1); plot.grid_rowconfigure(0, weight=1)

        self.fig = plt.figure(figsize=(10, 8))
        face = '#1e1e1e' if ctk.get_appearance_mode()=="Dark" else "#ffffff"
        self.fig.patch.set_facecolor(face)
        self.ax_s11 = self.fig.add_subplot(2,1,1); self.ax_s11.set_facecolor(face)
        self.ax_ff  = self.fig.add_subplot(2,1,2); self.ax_ff.set_facecolor(face)
        for ax in [self.ax_s11, self.ax_ff]: ax.grid(True, alpha=0.35)

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def _tab_log(self):
        tab = self.tabs.tab("Log")
        tab.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(tab, text="Log", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.log_box = ctk.CTkTextbox(tab, width=1000, height=500, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self._log(f"Log iniciado em {datetime.now():%Y-%m-%d %H:%M:%S}")

    # ---------- Log helpers ----------
    def _log(self, msg): self.log_q.put(f"[{datetime.now():%H:%M:%S}] {msg}")
    def _log_pump(self):
        try:
            while True:
                m = self.log_q.get_nowait()
                self.log_box.insert("end", m+"\n"); self.log_box.see("end")
        except queue.Empty:
            pass
        self.win.after(120, self._log_pump)

    # ---------- GUI actions ----------
    def _get_params(self) -> bool:
        try:
            def f(k): return float(getattr(self, f"ent_{k}").get())
            self.p["frequency"]   = f("frequency")
            self.p["gain"]        = f("gain")
            self.p["sweep_start"] = f("sweep_start")
            self.p["sweep_stop"]  = f("sweep_stop")
            self.p["er"]          = f("er")
            self.p["tan_d"]       = f("tan_d")
            self.p["h_sub"]       = f("h_sub")
            self.p["t_cu"]        = f("t_cu")
            self.p["probe_radius"]= f("probe_radius")
            self.p["coax_ba_ratio"]=f("coax_ba_ratio")
            self.p["coax_wall"]   = f("coax_wall")
            self.p["coax_Lp"]     = f("coax_Lp")
            self.p["antipad"]     = f("antipad")
            self.p["ovl"]         = f("ovl")
            self.p["sweep_step"]  = f("sweep_step")
            self.p["cores"]       = int(float(self.ent_cores.get()))
            self.p["substrate_material"] = self.cb_mat.get()
            self.p["sweep_type"]  = self.cb_sweep.get()
            self.p["spacing_scale"]= float(self.cb_spacing.get())
            return True
        except Exception as e:
            self.lb_status.configure(text=f"Erro: {e}")
            self._log(f"Erro parâmetros: {e}")
            return False

    def _suggest_mn(self):
        if not self._get_params(): return
        r, c, n = suggest_rows_cols_from_gain(self.p["gain"])
        self.lb_sug.configure(text=f"Sugestão: {r} × {c}  ({n} elementos)")
        self.sl_rows.set(r); self.lb_rows.configure(text=str(r))
        self.sl_cols.set(c); self.lb_cols.configure(text=str(c))

    def _apply_mn(self):
        self.rows = int(round(self.sl_rows.get()))
        self.cols = int(round(self.sl_cols.get()))
        self.n_total = self.rows * self.cols
        self._log(f"Layout aplicado: {self.rows} × {self.cols} = {self.n_total}")

    def _calculate(self):
        if not self._get_params(): return
        # Dimensões patch e espaçamento
        L_mm, W_mm, lamg50 = patch_dims_er(self.p["frequency"], self.p["er"], self.p["h_sub"])
        self.calc["patchL"] = L_mm; self.calc["patchW"] = W_mm; self.calc["lamg50"] = lamg50

        lam0_mm = C0 / (self.p["frequency"]*1e9) * 1000.0
        s = self.p["spacing_scale"] * lam0_mm
        self.calc["spacingX"] = s; self.calc["spacingY"] = s

        total_w = self.cols * W_mm + (self.cols - 1) * s
        total_l = self.rows * L_mm + (self.rows - 1) * s
        margin = self.p["margin_factor"] * max(total_w, total_l)
        self.calc["subW"] = total_w + 2*margin
        self.calc["subL"] = total_l + 2*margin

        # Larguras de linha
        h = self.p["h_sub"]; er = self.p["er"]
        self.calc["W50"]  = hammerstad_w(er, h, 50.0)
        self.calc["W70"]  = hammerstad_w(er, h, 70.710678)
        self.calc["W100"] = hammerstad_w(er, h, 100.0)
        self.calc["W141"] = hammerstad_w(er, h, 141.421356)
        self.calc["W200"] = hammerstad_w(er, h, 200.0)

        # λ/4 de 70,7 e 141 (para variável global)
        self.calc["Lq70"]  = guided_lambda_mm(self.p["frequency"], er, h, self.calc["W70"]) / 4.0
        self.calc["Lq141"] = guided_lambda_mm(self.p["frequency"], er, h, self.calc["W141"]) / 4.0

        self.lb_dims.configure(text=(f"Patch: {L_mm:.2f}×{W_mm:.2f} mm | λg(50Ω): {lamg50:.2f} mm | "
                                     f"Substrato: {self.calc['subW']:.1f}×{self.calc['subL']:.1f} mm"))
        self.lb_status.configure(text="Parâmetros calculados.")
        self._log("Parâmetros e variáveis calculados.")

    def _save_params(self):
        d = {"user": self.p, "calc": self.calc, "rows": self.rows, "cols": self.cols}
        with open("antenna_parameters.json", "w") as f: json.dump(d, f, indent=2)
        self._log("Parâmetros salvos.")

    def _load_params(self):
        try:
            with open("antenna_parameters.json") as f: d = json.load(f)
            self.p.update(d.get("user", {}))
            self.calc.update(d.get("calc", {}))
            self.rows = d.get("rows", self.rows)
            self.cols = d.get("cols", self.cols)
            # refletir
            for k,v in self.p.items():
                e = getattr(self, f"ent_{k}", None)
                if isinstance(e, ctk.CTkEntry):
                    e.delete(0,"end"); e.insert(0,str(v))
            self.cb_mat.set(self.p["substrate_material"])
            self.cb_sweep.set(self.p["sweep_type"])
            self.cb_spacing.set(str(self.p["spacing_scale"]))
            self.sl_rows.set(self.rows); self.lb_rows.configure(text=str(self.rows))
            self.sl_cols.set(self.cols); self.lb_cols.configure(text=str(self.cols))
            self._log("Parâmetros carregados.")
        except Exception as e:
            self._log(f"Erro ao carregar: {e}")

    # ---------- AEDT helpers ----------
    def _safe_name(self, base: str) -> str:
        self._id += 1
        base = re.sub(r"[^A-Za-z0-9_]", "_", base)
        return f"{base}_{self._id:04d}"

    def _set_var(self, name: str, value_mm: float):
        self.hfss[name] = f"{float(value_mm):.6f}mm"

    def _set_unitless(self, name: str, value: float):
        self.hfss[name] = float(value)

    # --- geometrias no plano z=h_sub ---
    def _rect_xy(self, x, y, w, h, name, mat="copper"):
        # aceita w/h como float ou nome de variável (str)
        if isinstance(w, str): wsize = w
        else: wsize = float(w)
        if isinstance(h, str): hsize = h
        else: hsize = float(h)
        return self.hfss.modeler.create_rectangle("XY", [x, y, "h_sub"], [wsize, hsize],
                                                  name=self._safe_name(name), material=mat)

    def _rect_h_run(self, x1, x2, y, wvar, name):
        """Retângulo horizontal (wvar = nome de variável da largura)."""
        x_min, x_max = (x1, x2) if x1 <= x2 else (x2, x1)
        return self._rect_xy(x_min, y - f"({wvar})/2", x_max - x_min, wvar, name)

    def _rect_v_run(self, y1, y2, x, wvar, name):
        y_min, y_max = (y1, y2) if y1 <= y2 else (y2, y1)
        return self._rect_xy(x - f"({wvar})/2", y_min, wvar, y_max - y_min, name)

    # --- portas/ coaxes ---
    def _coax_lumped(self, x0, y0, ground, substrate):
        p = self.p
        a = p["probe_radius"]; b = max(a * p["coax_ba_ratio"], a + 0.02)
        wall = p["coax_wall"]; Lp = p["coax_Lp"]; h = p["h_sub"]; clr=p["antipad"]

        # interno
        pin = self.hfss.modeler.create_cylinder("Z", [x0, y0, -Lp], a, h + Lp + 1e-3,
                                                name=self._safe_name("P0_Pin"), material="copper")
        # shield (só abaixo do GND)
        sh_o = self.hfss.modeler.create_cylinder("Z", [x0, y0, -Lp], b + wall, Lp,
                                                 name=self._safe_name("P0_Shield"), material="copper")
        sh_v = self.hfss.modeler.create_cylinder("Z", [x0, y0, -Lp], b, Lp,
                                                 name=self._safe_name("P0_ShieldVoid"), material="vacuum")
        self.hfss.modeler.subtract(sh_o, [sh_v], keep_originals=False)

        # antipad e furo
        hole_r = b + clr
        sub_hole = self.hfss.modeler.create_cylinder("Z", [x0, y0, 0.0], hole_r, h,
                                                     name=self._safe_name("P0_SubHole"), material="vacuum")
        try: self.hfss.modeler.subtract(substrate, [sub_hole], keep_originals=False)
        except Exception as e: self._log(f"Antipad sub aviso: {e}")
        g_hole = self.hfss.modeler.create_circle("XY", [x0, y0, 0.0], hole_r,
                                                 name=self._safe_name("P0_GHole"), material="vacuum")
        try: self.hfss.modeler.subtract(ground, [g_hole], keep_originals=False)
        except Exception as e: self._log(f"Antipad gnd aviso: {e}")

        # anel do porto
        ring = self.hfss.modeler.create_circle("XY", [x0, y0, -Lp], b,
                                               name=self._safe_name("P0_Ring"), material="vacuum")
        hole = self.hfss.modeler.create_circle("XY", [x0, y0, -Lp], a,
                                               name=self._safe_name("P0_RingHole"), material="vacuum")
        self.hfss.modeler.subtract(ring, [hole], keep_originals=False)

        # integração radial
        eps = min(0.1*(b-a), 0.05)
        p1 = [x0 + a + eps, y0, -Lp]; p2 = [x0 + b - eps, y0, -Lp]
        self.hfss.lumped_port(assignment=ring.name, integration_line=[p1, p2],
                              impedance=50.0, name=self._portname, renormalize=True)
        # pad no topo e curtíssimo 50 Ω até (0,0) para casado visual
        top = self.hfss.modeler.create_circle("XY", [x0, y0, "h_sub"], a,
                                              name=self._safe_name("P0_TopPad"), material="copper")
        return top

    # ---------- H-tree (corporate) ----------
    def _qwl(self, z0: float) -> float:
        return guided_lambda_mm(self.p["frequency"], self.p["er"], self.p["h_sub"],
                                hammerstad_w(self.p["er"], self.p["h_sub"], z0)) / 4.0

    def _build_x_tree(self, xs: List[float], y: float, parent_x: float) -> List[Tuple[float,float]]:
        """
        H-tree horizontal: da junção (parent_x,y) até colunas xs.
        Retorna pontos de junção por coluna [(x_col, y)].
        """
        xs = sorted(xs)
        leaves: List[Tuple[float,float]] = []
        if len(xs) == 1:
            # ligação direta: QW70 + 100Ω até x_col
            x_col = xs[0]
            self._rect_h_run(parent_x, parent_x + math.copysign(self.calc["Lq70"], x_col - parent_x),
                             y, "W70", "X_Q70")
            self._rect_h_run(parent_x + math.copysign(self.calc["Lq70"], x_col - parent_x), x_col, y, "W100", "X_100")
            return [(x_col, y)]

        # split em dois grupos
        mid = len(xs)//2
        left, right = xs[:mid], xs[mid:]
        # junções dos filhos no meio de cada grupo
        xL = 0.5*(left[0] + left[-1])
        xR = 0.5*(right[0] + right[-1])

        # parent -> filho L
        dirL = math.copysign(1.0, xL - parent_x)
        self._rect_h_run(parent_x, parent_x + dirL*self.calc["Lq70"], y, "W70", "X_Q70_L")
        self._rect_h_run(parent_x + dirL*self.calc["Lq70"], xL, y, "W100", "X_100_L")
        # parent -> filho R
        dirR = math.copysign(1.0, xR - parent_x)
        self._rect_h_run(parent_x, parent_x + dirR*self.calc["Lq70"], y, "W70", "X_Q70_R")
        self._rect_h_run(parent_x + dirR*self.calc["Lq70"], xR, y, "W100", "X_100_R")

        leaves += self._build_x_tree(left, y, xL)
        leaves += self._build_x_tree(right, y, xR)
        return leaves

    def _build_y_tree(self, x: float, ys: List[float], parent_y: float) -> None:
        """
        H-tree vertical: da junção (x,parent_y) até linhas ys (centros dos patches).
        Na folha: λ/4 141Ω e stub 200Ω horizontal até o patch.
        """
        ys = sorted(ys)
        if len(ys) == 1:
            y_p = ys[0]
            # do parent à proximidade do patch (100Ω) + QW141 vertical
            dirU = math.copysign(1.0, y_p - parent_y)
            # 100Ω até aproximar Lq141/2 antes do patch
            run = max(0.2, abs(y_p - parent_y) - self.calc["Lq141"]/2 - 0.2)
            y1 = parent_y
            y2 = parent_y + dirU*run
            if run > 0:
                self._rect_v_run(y1, y2, x, "W100", "Y_100")
            # QW141
            self._rect_v_run(y2, y2 + dirU*self.calc["Lq141"], x, "W141", "Y_Q141")
            # stub 200Ω horizontal até a borda do patch (borda interna)
            # (stub criado junto do patch na geração dos patches, OVL garante contato)
            return

        # split
        mid = len(ys)//2
        low, high = ys[:mid], ys[mid:]
        yL = 0.5*(low[0] + low[-1])
        yH = 0.5*(high[0] + high[-1])

        # parent -> filho baixo
        dirD = math.copysign(1.0, yL - parent_y)
        self._rect_v_run(parent_y, parent_y + dirD*self.calc["Lq70"], x, "W70", "Y_Q70_L")
        self._rect_v_run(parent_y + dirD*self.calc["Lq70"], yL, x, "W100", "Y_100_L")
        # parent -> filho alto
        dirU = math.copysign(1.0, yH - parent_y)
        self._rect_v_run(parent_y, parent_y + dirU*self.calc["Lq70"], x, "W70", "Y_Q70_H")
        self._rect_v_run(parent_y + dirU*self.calc["Lq70"], yH, x, "W100", "Y_100_H")

        self._build_y_tree(x, low, yL)
        self._build_y_tree(x, high, yH)

    # ---------- Simulação ----------
    def _start(self):
        if self.sim_running:
            self._log("Simulação já em execução."); return
        if not self._get_params(): return
        if "patchL" not in self.calc: self._calculate()

        self.sim_running = True
        self.bt_run.configure(state="disabled"); self.bt_stop.configure(state="normal")
        self.pb.set(0.0); self.lb_sim.configure(text="Iniciando…")

        threading.Thread(target=self._run_sim, daemon=True).start()

    def _stop(self):
        self._log("Parada solicitada (não hard-kill).")
        self.sim_running = False

    def _run_sim(self):
        try:
            self.lb_sim.configure(text="Abrindo AEDT…"); self.pb.set(0.05)
            # Desktop
            if self.desktop is None:
                self.desktop = Desktop(version="2024.2", non_graphical=False, new_desktop=True)
                self._log("Desktop inicializado.")

            if self.tempdir is None:
                self.tempdir = tempfile.TemporaryDirectory(suffix=".ansys")
            proj = os.path.join(self.tempdir.name, f"patch_array_{datetime.now():%Y%m%d_%H%M%S}.aedt")

            oDesktop = self.desktop.odesktop
            oDesktop.NewProject()
            oProject = oDesktop.GetActiveProject()
            if "patch_array" not in [d.GetName() for d in oProject.GetDesigns()]:
                oProject.InsertDesign("HFSS", "patch_array", "DrivenModal", "")
            oProject.SetActiveDesign("patch_array")
            try: oProject.SaveAs(proj, True)
            except Exception: pass

            self.hfss = Hfss(project=oProject.GetName(), design="patch_array",
                             solution_type="DrivenModal", new_desktop=False)
            self.hfss.modeler.model_units = "mm"
            self._log(f"Projeto ativo: {proj}")

            # Materiais
            mat = self.p["substrate_material"]
            if not self.hfss.materials.checkifmaterialexists(mat):
                mat = "Custom_Substrate"
                if not self.hfss.materials.checkifmaterialexists(mat):
                    self.hfss.materials.add_material(mat)
                m = self.hfss.materials.material_keys[mat]
                m.permittivity = self.p["er"]; m.dielectric_loss_tangent = self.p["tan_d"]

            # ---- Variáveis de design (PARÂMETROS) ----
            # Tecnológicos / geom. macro
            self._set_var("h_sub", self.p["h_sub"])
            self._set_var("t_cu",  self.p["t_cu"])
            self._set_var("patchL", self.calc["patchL"])
            self._set_var("patchW", self.calc["patchW"])
            self._set_var("spacingX", self.calc["spacingX"])
            self._set_var("spacingY", self.calc["spacingY"])
            self._set_var("subW", self.calc["subW"])
            self._set_var("subL", self.calc["subL"])
            self._set_var("W50",  self.calc["W50"])
            self._set_var("W70",  self.calc["W70"])
            self._set_var("W100", self.calc["W100"])
            self._set_var("W141", self.calc["W141"])
            self._set_var("W200", self.calc["W200"])
            self._set_var("Lq70",  self.calc["Lq70"])
            self._set_var("Lq141", self.calc["Lq141"])
            self._set_var("OVL",   self.p["ovl"])

            # ---- Substrato e GND (usando VARIÁVEIS) ----
            substrate = self.hfss.modeler.create_box(
                ["-subW/2","-subL/2", 0], ["subW","subL","h_sub"],
                name=self._safe_name("Substrate"), material=mat)
            ground = self.hfss.modeler.create_rectangle(
                "XY", ["-subW/2","-subL/2", 0], ["subW","subL"],
                name=self._safe_name("Ground"), material="copper")

            # ---- Porta coaxial no centro (0,0) ----
            top_pad = self._coax_lumped(0.0, 0.0, ground, substrate)

            # ---- Patches + stub 200 Ω parametrizados ----
            Wp = self.calc["patchW"]; Lp = self.calc["patchL"]
            sx = self.calc["spacingX"]; sy = self.calc["spacingY"]

            xs = [-( (self.cols-1)/2.0 - j)*(Wp + sx) for j in range(self.cols)]
            ys = [ ( (self.rows-1)/2.0 - i)*(Lp + sy) for i in range(self.rows)]

            # cria patches centrados em (x,y) e um stub 200 Ω OBLIGATORIAMENTE conectado
            for x in xs:
                for y in ys:
                    # patch com variáveis patchW/patchL
                    self.hfss.modeler.create_rectangle("XY",
                        [f"{x}-patchW/2", f"{y}-patchL/2", "h_sub"],
                        ["patchW", "patchL"],
                        name=self._safe_name("Patch"), material="copper")

                    # stub 200 Ω horizontal do eixo da coluna (x) à borda interna do patch
                    # lado depende de x sinal (para centralizar a coluna no centro do patch)
                    if x >= 0:
                        # borda esquerda do patch: inicia um pouco antes para OVL
                        x_edge = x - Wp/2.0
                        self._rect_xy(f"{x_edge}-OVL", f"{y}-W200/2", f"{abs(x - x_edge)+self.p['ovl']:.6f}", "W200", "STUB200_L")
                    else:
                        # borda direita
                        x_edge = x + Wp/2.0
                        self._rect_xy(f"{x}", f"{y}-W200/2", f"{abs(x_edge - x)+self.p['ovl']:.6f}", "W200", "STUB200_R")

            # ---- Tronco central 50 Ω curto (pad -> rede) ----
            self._rect_xy("-W50/2", "-W50/2", "W50", "W50", "FEED50")  # quadradinho de 50Ω sobre o topo do pad

            # ---- H-tree: X (para colunas), depois Y (por coluna) ----
            self._log("Construindo rede corporativa…")
            leaves = self._build_x_tree(xs, 0.0, 0.0)
            for x_col, _ in leaves:
                ys_col = sorted(ys)
                self._build_y_tree(x_col, ys_col, 0.0)

            # ---- Boundaries ----
            self._log("Definindo PerfectE…")
            names = [ground.name]
            try:
                if hasattr(top_pad, "name"):
                    names.append(top_pad.name)
            except Exception:
                pass
            # não tentamos unir; muitas folhas podem falhar; PerfectE cobre as folhas top metal
            try:
                self.hfss.assign_perfecte_to_sheets(list(dict.fromkeys(names)))
            except Exception as e:
                self._log(f"Aviso PerfectE: {e}")

            # ---- Região + radiação ----
            self._log("Criando região e Radiação…")
            lam0_start_mm = C0 / (self.p["sweep_start"] * 1e9) * 1000.0
            pad_mm = max(5.0, min(0.25 * lam0_start_mm, 15.0))
            region = self.hfss.modeler.create_region([pad_mm]*6, is_percentage=False)
            self.hfss.assign_radiation_boundary_to_objects(region)

            # Infinite sphere
            try:
                rf = self.hfss.odesign.GetModule("RadField")
                rf.InsertInfiniteSphereSetup([
                    "NAME:IS1","UseCustomRadiationSurface:=",False,"CSDefinition:=","Theta-Phi","Polarization:=","Linear",
                    "ThetaStart:=","-180deg","ThetaStop:=","180deg","ThetaStep:=","2deg",
                    "PhiStart:=","-180deg","PhiStop:=","180deg","PhiStep:=","2deg","UseLocalCS:=",False])
            except Exception as e:
                self._log(f"InfiniteSphere aviso: {e}")

            self.pb.set(0.70); self.lb_sim.configure(text="Configurando Setup…")

            # ---- Setup + sweep ----
            setup = self.hfss.create_setup(name="Setup1", setup_type="HFSSDriven")
            setup.props["Frequency"] = f"{self.p['frequency']}GHz"
            setup.props["MaxDeltaS"] = 0.02
            st = self.p["sweep_type"]
            try:
                if st == "Discrete":
                    setup.create_linear_step_sweep(unit="GHz", start_frequency=self.p["sweep_start"],
                                                   stop_frequency=self.p["sweep_stop"], step_size=self.p["sweep_step"], name="Sweep1")
                elif st == "Fast":
                    setup.create_frequency_sweep(unit="GHz", name="Sweep1",
                                                 start_frequency=self.p["sweep_start"],
                                                 stop_frequency=self.p["sweep_stop"], sweep_type="Fast")
                else:
                    setup.create_frequency_sweep(unit="GHz", name="Sweep1",
                                                 start_frequency=self.p["sweep_start"],
                                                 stop_frequency=self.p["sweep_stop"], sweep_type="Interpolating")
            except Exception as e:
                self._log(f"Sweep aviso: {e}")

            if bool(self.ch_save.get()):
                self.hfss.save_project()

            self.lb_sim.configure(text="Analisando…"); self.pb.set(0.78)
            self.hfss.analyze_setup("Setup1", cores=int(self.p.get("cores", 4)))

            # ---- Resultados ----
            self.lb_sim.configure(text="Extraindo resultados…"); self.pb.set(0.92)
            self._plot_results()
            self.lb_sim.configure(text="Concluído"); self.pb.set(1.0)
            self._log("Simulação concluída.")
        except Exception as e:
            self._log(f"ERRO: {e}\n{traceback.format_exc()}")
            self.lb_sim.configure(text=f"Erro: {e}")
        finally:
            self.bt_run.configure(state="normal"); self.bt_stop.configure(state="disabled")
            self.sim_running = False

    # ---------- Resultados ----------
    def _plot_results(self):
        self._log("Plotando…")
        self.ax_s11.clear(); self.ax_ff.clear()
        self.ax_s11.grid(True, alpha=0.35); self.ax_ff.grid(True, alpha=0.35)
        self.ax_s11.set_title("S11"); self.ax_s11.set_xlabel("Frequência (GHz)"); self.ax_s11.set_ylabel("dB")
        self.ax_ff.set_title(f"Ganho — cortes @ {self.p['frequency']} GHz"); self.ax_ff.set_xlabel("Ângulo (graus)"); self.ax_ff.set_ylabel("dB")

        # --- S11 robusto ---
        sol = None
        port_exprs = [f"dB(S({self._portname},{self._portname}))", "dB(S(1,1))"]
        ctxs = ["Setup1: Sweep1", "Setup1 : Sweep1", "Setup1 : LastAdaptive", "Setup1:LastAdaptive"]
        try:
            for expr in port_exprs:
                try:
                    rpt = self.hfss.post.reports_by_category.standard(expressions=[expr])
                    ok = False
                    for ctx in ctxs:
                        try:
                            rpt.context = [ctx]
                            sol = rpt.get_solution_data()
                            if sol and hasattr(sol, "primary_sweep_values"):
                                ok = True; break
                        except Exception:
                            sol = None
                    if ok: break
                except Exception:
                    continue

            if sol:
                f = np.asarray(sol.primary_sweep_values, dtype=float)
                dat = sol.data_real()
                y = np.asarray(dat[0] if isinstance(dat, (list,tuple)) else dat, dtype=float)
                if f.size and f.size == y.size:
                    self.sim_data = np.column_stack((f, y))
                    self.ax_s11.plot(f, y, lw=2, label="S11")
                    self.ax_s11.axhline(-10, ls="--", alpha=0.6, label="-10 dB")
                    self.ax_s11.axvline(self.p["frequency"], ls="--", alpha=0.6)
                    self.ax_s11.legend()
                else:
                    self.ax_s11.text(0.5,0.5,"S11 indisponível",ha="center",va="center",transform=self.ax_s11.transAxes)
            else:
                self.ax_s11.text(0.5,0.5,"S11 indisponível",ha="center",va="center",transform=self.ax_s11.transAxes)
        except Exception as e:
            self._log(f"S11 falhou: {e}")
            self.ax_s11.text(0.5,0.5,"S11 indisponível",ha="center",va="center",transform=self.ax_s11.transAxes)

        # --- Far-field (dois cortes) robusto ---
        try:
            ff_ok = False
            # θ (Phi=0)
            rep1 = self.hfss.post.reports_by_category.far_field(
                expressions=["dB(GainTotal)"], context="IS1", primary_sweep_variable="Theta",
                setup="Setup1 : LastAdaptive",
                variations={"Freq": f"{self.p['frequency']}GHz", "Phi": "0deg", "Theta": "All"},
                name="FF_Theta_Phi0")
            sd1 = rep1.get_solution_data() if hasattr(rep1, "get_solution_data") else None
            if sd1:
                th = np.asarray(sd1.primary_sweep_values, dtype=float)
                g1 = np.asarray(sd1.data_real()[0], dtype=float)
                if th.size > 1 and th.size == g1.size:
                    self.ax_ff.plot(th, g1, lw=2, label="Phi=0°"); ff_ok = True

            # φ (Theta=90)
            rep2 = self.hfss.post.reports_by_category.far_field(
                expressions=["dB(GainTotal)"], context="IS1", primary_sweep_variable="Phi",
                setup="Setup1 : LastAdaptive",
                variations={"Freq": f"{self.p['frequency']}GHz", "Theta": "90deg", "Phi": "All"},
                name="FF_Phi_Theta90")
            sd2 = rep2.get_solution_data() if hasattr(rep2, "get_solution_data") else None
            if sd2:
                ph = np.asarray(sd2.primary_sweep_values, dtype=float)
                g2 = np.asarray(sd2.data_real()[0], dtype=float)
                if ph.size > 1 and ph.size == g2.size:
                    self.ax_ff.plot(ph, g2, lw=2, label="Theta=90°")

            if ff_ok: self.ax_ff.legend()
            else:
                self.ax_ff.text(0.5,0.5,"Far-field indisponível",ha="center",va="center",transform=self.ax_ff.transAxes)
        except Exception as e:
            self._log(f"FF falhou: {e}")
            self.ax_ff.text(0.5,0.5,"Far-field indisponível",ha="center",va="center",transform=self.ax_ff.transAxes)

        self.fig.tight_layout(); self.canvas.draw()
        self._log("Plot OK.")

    # ---------- Encerramento ----------
    def _close(self):
        self._log("Fechando…")
        try:
            if self.hfss:
                try:
                    if bool(self.ch_save.get()): self.hfss.save_project()
                except Exception as e: self._log(f"Salvar aviso: {e}")
            if self.desktop:
                try:
                    self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception as e: self._log(f"Release aviso: {e}")
            if self.tempdir: 
                try: self.tempdir.cleanup()
                except Exception: pass
        finally:
            try: self.win.destroy()
            except Exception: pass

# ---------- Run ----------
if __name__ == "__main__":
    app = PatchArrayApp()
    app.win.mainloop()
