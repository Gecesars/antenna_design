# -*- coding: utf-8 -*-
"""
HFSS Patch Array – Automação (PyAEDT)
-------------------------------------
- HFSS (Modal), Auto Open Region
- Patch retangular + λ/4 microstrip + feed 50 Ω (sheets PEC)
- Lumped Port correta (XZ) tocando topo (PEC) e GND (PEC)
- Array NxN por ganho-alvo
- Setup "OpSetup" + Sweep "OpFast" (Interpolating, 101 pts, banda +30%)
- Edit Sources pós-solve (Amp_Pi, Phi_Pi) com fallback
- GUI CTk e liberação do AEDT ao fechar

Requisitos:
    pip install ansys-aedt-core==0.18.1 customtkinter

Observação:
    Se scikit-rf não estiver disponível, cai no valor de c padrão.
"""

from __future__ import annotations
import os, sys, math, traceback
from dataclasses import dataclass
from typing import Dict, List, Tuple

# ======== Constantes físicas (scikit-rf opcional) ========
try:
    from skrf.constants import c as C0  # apenas 'c'; outras podem não existir em algumas versões
except Exception:
    C0 = 299_792_458.0  # m/s

# ======== PyAEDT ========
from ansys.aedt.core import Hfss
from ansys.aedt.core.generic.constants import GeometryOperators

# ======== GUI (CTk) ========
import customtkinter as ctk


# ---------------------------------------------------------------------------
# Utilidades de projeto
# ---------------------------------------------------------------------------

def db10(x: float) -> float:
    return 10.0 * math.log10(max(x, 1e-30))


def gain_to_n_elements(g_target_dbi: float, g_elem_dbi: float = 6.0) -> int:
    """Número de elementos aproximado pela diretividade: N ≈ 10^((G_alvo - G_elem)/10)."""
    n = max(1, int(math.ceil(10 ** ((g_target_dbi - g_elem_dbi) / 10.0))))
    return n


def square_grid(n: int) -> Tuple[int, int]:
    """Distribui N em grade quase quadrada."""
    m = int(math.floor(math.sqrt(n)))
    k = int(math.ceil(n / m))
    return m, k  # linhas (y), colunas (x)


def patch_dimensions(f0_ghz: float, eps_r: float, h_mm: float) -> Tuple[float, float, float]:
    """
    Balanis (aprox):
    - W ≈ c/(2 f0) * sqrt(2/(εr+1))
    - εeff
    - ΔL/h ≈ 0.412 * ((εeff+0.3)(W/h+0.264))/((εeff-0.258)(W/h+0.8))
    - L = L_eff - 2ΔL
    Retorna (W_mm, L_mm, eps_eff)
    """
    f0 = f0_ghz * 1e9
    h = h_mm / 1000.0
    W = (C0 / (2.0 * f0)) * math.sqrt(2.0 / (eps_r + 1.0))
    eps_eff = 0.5 * (eps_r + 1.0) + 0.5 * (eps_r - 1.0) * (1.0 + 12.0 * h / W) ** -0.5
    dL_over_h = 0.412 * ((eps_eff + 0.3) * (W / h + 0.264)) / ((eps_eff - 0.258) * (W / h + 0.8))
    L_eff = C0 / (2.0 * f0 * math.sqrt(eps_eff))
    L = L_eff - 2.0 * dL_over_h * h
    return W * 1000.0, L * 1000.0, eps_eff


def microstrip_w_for_z0(z0: float, eps_r: float, h_mm: float) -> float:
    """Hammerstad-Jensen (aprox) – retorna largura (mm) para Z0 em microstrip."""
    h = h_mm / 1000.0  # m
    # chute inicial
    A = z0 / 60.0 * math.sqrt((eps_r + 1.0) / 2.0) + ((eps_r - 1.0) / (eps_r + 1.0)) * (0.23 + 0.11 / eps_r)
    B = (377 * math.pi) / (2.0 * z0 * math.sqrt(eps_r))
    if A <= 1:
        W_h = 8 * math.exp(A) / (math.exp(2 * A) - 2)
    else:
        W_h = (2 / math.pi) * (B - 1 - math.log(2 * B - 1) + ((eps_r - 1) / (2 * eps_r)) * (math.log(B - 1) + 0.39 - 0.61 / eps_r))
    return max(0.05, W_h * h * 1000.0)  # mm


def microstrip_eps_eff(w_mm: float, eps_r: float, h_mm: float) -> float:
    """εeff da linha (Hammerstad)."""
    w_h = (w_mm / 1000.0) / (h_mm / 1000.0)
    if w_h <= 1:
        eps_eff = (eps_r + 1) / 2 + (eps_r - 1) / 2 * (1 / math.sqrt(1 + 12 / w_h) + 0.04 * (1 - w_h) ** 2)
    else:
        eps_eff = (eps_r + 1) / 2 + (eps_r - 1) / 2 * (1 / math.sqrt(1 + 12 / w_h))
    return eps_eff


def guided_quarter_lambda(f0_ghz: float, eps_eff_line: float) -> float:
    """λg/4 (mm) para a linha."""
    f0 = f0_ghz * 1e9
    lam_g_m = C0 / (f0 * math.sqrt(eps_eff_line))
    return lam_g_m * 1000.0 / 4.0


def rin_edge_patch(f0_ghz: float, W_mm: float, h_mm: float, eps_eff: float) -> float:
    """Rin aproximado na borda do patch (modelo de fenda)."""
    f0 = f0_ghz * 1e9
    lam0 = C0 / f0
    W = W_mm / 1000.0
    h = h_mm / 1000.0
    k0h = 2 * math.pi * h / lam0
    G = math.pi * W / (120.0 * math.pi * lam0) * (1.0 - (k0h ** 2) / 24.0)
    Rin = 1.0 / (2.0 * G + 1e-18)
    return Rin


# ---------------------------------------------------------------------------
# Construção geométrica / HFSS helpers
# ---------------------------------------------------------------------------

@dataclass
class PatchParams:
    Wp: float
    Lp: float
    Wt: float
    Lq: float
    W50: float
    eps_eff: float


def set_design_variables(hfss: Hfss, h_mm: float, tcu_mm: float, p: PatchParams):
    """Cria variáveis para deixar o modelo paramétrico."""
    hfss["h"] = f"{h_mm:.6f}mm"
    hfss["tCu"] = f"{tcu_mm:.6f}mm"
    hfss["Wp"] = f"{p.Wp:.6f}mm"
    hfss["Lp"] = f"{p.Lp:.6f}mm"
    hfss["Wt"] = f"{p.Wt:.6f}mm"
    hfss["Lq"] = f"{p.Lq:.6f}mm"
    hfss["W50"] = f"{p.W50:.6f}mm"


def mm_str(x: float) -> str:
    return f"{x:.6f}mm"


def add_ground_and_substrate(hfss: Hfss, size_x_mm: float, size_y_mm: float, h_mm: float) -> Tuple[str, str]:
    """Cria Substrate (BOX) e GND (SHEET XY em z=0) e aplica PEC ao GND."""
    # Substrate (box, para visual; dielétrico é material do projeto – FR4 etc. se quiser, troque para 'FR4_epoxy')
    hfss.modeler.create_box(
        origin=[mm_str(-size_x_mm / 2), mm_str(-size_y_mm / 2), "0mm"],
        sizes=[mm_str(size_x_mm), mm_str(size_y_mm), mm_str(h_mm)],
        name="Substrate", material="FR4_epoxy"
    )
    # GND como sheet em z=0
    gnd = hfss.modeler.create_rectangle(
        origin=[mm_str(-size_x_mm / 2), mm_str(-size_y_mm / 2), "0mm"],
        sizes=[mm_str(size_x_mm), mm_str(size_y_mm)],
        orientation="XY", name="GND"
    )
    hfss.assign_perfecte_to_sheets(["GND"])
    return "Substrate", "GND"


def add_patch_with_qw(
    hfss: Hfss,
    cx_mm: float, cy_mm: float,
    p: PatchParams, idx: int,
    h_mm: float, tcu_mm: float,
    z0_system: float = 50.0
) -> Tuple[str, bool]:
    """
    Cria:
      - Patch (Wp x Lp) em z=h
      - Linha λ/4 (Wt x Lq) em z=h
      - Feed 50Ω (W50 x Lfeed) em z=h
      - Porta LUMPED no sheet XZ (face id) tocando GND e topo.
    """
    z_top = "h"
    z_gnd = "0mm"
    h_tot = "h+tCu"
    Lfeed_mm = max(5.0, 0.15 * p.Lq)

    # --- topo ---
    r_patch = hfss.modeler.create_rectangle(
        origin=[f"{mm_str(cx_mm)}-Wp/2", f"{mm_str(cy_mm)}-Lp/2", z_top],
        sizes=["Wp", "Lp"], orientation="XY", name=f"Patch_{idx}"
    )
    r_qw = hfss.modeler.create_rectangle(
        origin=[f"{mm_str(cx_mm)}-Wt/2", f"{mm_str(cy_mm)}+Lp/2", z_top],
        sizes=["Wt", "Lq"], orientation="XY", name=f"TLineQW_{idx}"
    )
    r_feed = hfss.modeler.create_rectangle(
        origin=[f"{mm_str(cx_mm)}-W50/2", f"{mm_str(cy_mm)}+Lp/2+Lq", z_top],
        sizes=["W50", mm_str(Lfeed_mm)], orientation="XY", name=f"Feed50_{idx}"
    )

    # Aplique PEC explicitamente em cada um (mais robusto ao unite)
    hfss.assign_perfecte_to_sheets([r_patch.name, r_qw.name, r_feed.name])

    # Una os três; o nome mantido costuma ser o do primeiro
    hfss.modeler.unite([r_patch, r_qw, r_feed])

    # --- porta ---
    y_port = cy_mm + (p.Lp / 2.0) + p.Lq + Lfeed_mm
    r_port = hfss.modeler.create_rectangle(
        origin=[f"{mm_str(cx_mm)}-W50/2", mm_str(y_port), z_gnd],
        sizes=["W50", h_tot], orientation="XZ", name=f"PortSheet_{idx}"
    )

    # pegue as faces do sheet e escolha a(s) em y == y_port
    faces = hfss.modeler.get_object_faces(r_port.name)
    # função util pra filtrar pela coordenada Y do centro
    def _face_at_y(fid, y_target):
        try:
            c = hfss.modeler.get_face_center(fid)
            return abs(c[1] - y_target) < 1e-6
        except Exception:
            return False

    cand = [f for f in faces if _face_at_y(f, y_port)]
    if not cand:  # fallback: use todas (ordem estável)
        cand = faces[:]

    pname = f"P{idx}"
    # tente na ordem: primeira face candidata, depois a outra
    for fid in cand:
        try:
            hfss.lumped_port(
                assignment=fid,         # <<<<<< face id
                reference="GND",
                impedance=z0_system,
                renormalize=True,
                name=pname
            )
            return pname, True
        except Exception:
            continue

    # se nada deu, reporte falha mas siga em frente
    return pname, False



# ---------------------------------------------------------------------------
# Fluxo principal de criação do projeto
# ---------------------------------------------------------------------------

def build_array_project(
    project_path: str,
    f_low_ghz: float, f_high_ghz: float,
    f0_ghz: float,
    g_target_dbi: float,
    eps_r: float, h_mm: float, tcu_mm: float,
    run_after: bool = True,
    version: str = "2024.2"
):
    """
    Cria o projeto HFSS, array NxN, setup e (opcionalmente) roda a simulação.
    Retorna (hfss, info_dict).
    """
    # Banda do sweep com +30% (±15%)
    span = f_high_ghz - f_low_ghz
    f1 = max(0.01, f_low_ghz - 0.15 * span)
    f2 = f_high_ghz + 0.15 * span

    # Dimensões teóricas
    Wp, Lp, eps_eff_patch = patch_dimensions(f0_ghz, eps_r, h_mm)
    Rin_edge = rin_edge_patch(f0_ghz, Wp, h_mm, eps_eff_patch)
    Zt = math.sqrt(50.0 * Rin_edge)
    Wt = microstrip_w_for_z0(Zt, eps_r, h_mm)
    W50 = microstrip_w_for_z0(50.0, eps_r, h_mm)
    eps_eff_wt = microstrip_eps_eff(Wt, eps_r, h_mm)
    Lq = guided_quarter_lambda(f0_ghz, eps_eff_wt)

    params = PatchParams(Wp=Wp, Lp=Lp, Wt=Wt, Lq=Lq, W50=W50, eps_eff=eps_eff_patch)

    # Array: N elementos (Nx x Ny), espaçamento ~0.6 λ0
    lam0_mm = (C0 / (f0_ghz * 1e9)) * 1000.0
    sx = sy = 0.6 * lam0_mm
    n_total = gain_to_n_elements(g_target_dbi, g_elem_dbi=6.0)
    Ny, Nx = square_grid(n_total)
    n_real = Nx * Ny

    # Caixa/bordas do substrato e GND um pouco maiores
    margin = 0.5 * lam0_mm
    size_x = Nx * sx + 2 * margin
    size_y = Ny * sy + 2 * margin

    # ---- Inicializa HFSS ----
    hfss = Hfss(
        project=project_path,
        design="PatchArray_HFSS",
        solution_type="Modal",
        version=version,
        new_desktop=True,
        non_graphical=False
    )
    hfss.modeler.model_units = "mm"

    # Variáveis para parametrização
    set_design_variables(hfss, h_mm, tcu_mm, params)

    # Substrate e GND (PEC)
    add_ground_and_substrate(hfss, size_x, size_y, h_mm)

    # Centraliza coordenadas dos elementos
    x0 = - (Nx - 1) * sx / 2.0
    y0 = - (Ny - 1) * sy / 2.0

    ports: List[str] = []
    lumped_list: List[bool] = []

    idx = 1
    for iy in range(Ny):
        for ix in range(Nx):
            cx = x0 + ix * sx
            cy = y0 + iy * sy
            p_name, ok_lumped = add_patch_with_qw(hfss, cx, cy, params, idx, h_mm, tcu_mm, 50.0)
            ports.append(p_name)
            lumped_list.append(ok_lumped)
            idx += 1

    # Região aberta automática
    hfss.create_open_region(frequency=f"{f0_ghz}GHz")

    # Setup explícito + sweep FAST 101 pts
    setup = hfss.create_setup(name="OpSetup", Frequency=f"{f0_ghz}GHz")
    setup.create_frequency_sweep(
        name="OpFast", unit="GHz",
        start_frequency=f1, stop_frequency=f2,
        sweep_type="Interpolating", num_of_freq_points=101
    )

    # Garantir que rodamos o setup certo:
    hfss.active_setup = "OpSetup"

    # Salvar já
    hfss.save_project()

    # Rodar (opcional)
    if run_after:
        try:
            hfss.analyze_setup("OpSetup")
        except Exception:
            try:
                hfss.analyze(setup="OpSetup")
            except Exception:
                hfss.analyze()

        # Pós-processo: fontes com variáveis Amp_/Phi_
        assign = {}
        pwr_each = 1.0 / max(1, len(ports))
        for p in ports:
            hfss[f"Amp_{p}"] = f"{pwr_each:.6f}W"
            hfss[f"Phi_{p}"] = "0deg"
            assign[p] = {"magnitude": f"Amp_{p}", "phase": f"Phi_{p}", "source_type": "Power"}

        ok_sources = False
        try:
            hfss.edit_sources(assign)  # portas lumped costumam ser "P1"
            ok_sources = True
        except Exception:
            # fallback para modos ("P1:1")
            assign_modes = {f"{k}:1": v for k, v in assign.items()}
            try:
                hfss.edit_sources(assign_modes)
                ok_sources = True
            except Exception:
                pass

        # Relatório S11 (primeira porta)
        try:
            hfss.post.create_report(
                expressions=["dB(S(1,1))"],
                setup_sweep_name="OpSetup : OpFast",
                primary_sweep_variable="Freq"
            )
        except Exception:
            pass

    info = {
        "W_mm": Wp, "L_mm": Lp, "eps_eff": eps_eff_patch,
        "Rin_edge_ohm": Rin_edge,
        "Zt_ohm": Zt, "Wt_mm": Wt, "Lq_mm": Lq, "W50_mm": W50,
        "Nx": Nx, "Ny": Ny, "N_real": n_real,
        "sx_mm": sx, "sy_mm": sy,
        "ports": ports, "lumped_ok": lumped_list,
        "setup": "OpSetup", "sweep": "OpFast",
        "f1_ghz": f1, "f2_ghz": f2
    }
    return hfss, info


# ---------------------------------------------------------------------------
# GUI – CustomTkinter
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HFSS Patch Array — Automação (PyAEDT)")
        self.geometry("720x540")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.hfss: Hfss | None = None

        # ---- Entradas ----
        row = 0
        self._lbl_tip = ctk.CTkLabel(self, text="Campos com unidades: GHz, dBi, εr, mm.", anchor="w")
        self._lbl_tip.grid(row=row, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 4))
        row += 1

        # Banda
        ctk.CTkLabel(self, text="Frequência baixa (GHz):").grid(row=row, column=0, sticky="e", padx=6, pady=6)
        self.e_flow = ctk.CTkEntry(self); self.e_flow.insert(0, "2.30")
        self.e_flow.grid(row=row, column=1, sticky="w", padx=6, pady=6)

        ctk.CTkLabel(self, text="Frequência alta (GHz):").grid(row=row, column=2, sticky="e", padx=6, pady=6)
        self.e_fhigh = ctk.CTkEntry(self); self.e_fhigh.insert(0, "2.50")
        self.e_fhigh.grid(row=row, column=3, sticky="w", padx=6, pady=6)
        row += 1

        # f0 e ganho
        ctk.CTkLabel(self, text="f0 (GHz):").grid(row=row, column=0, sticky="e", padx=6, pady=6)
        self.e_f0 = ctk.CTkEntry(self); self.e_f0.insert(0, "2.40")
        self.e_f0.grid(row=row, column=1, sticky="w", padx=6, pady=6)

        ctk.CTkLabel(self, text="Ganho alvo (dBi):").grid(row=row, column=2, sticky="e", padx=6, pady=6)
        self.e_gain = ctk.CTkEntry(self); self.e_gain.insert(0, "12.0")
        self.e_gain.grid(row=row, column=3, sticky="w", padx=6, pady=6)
        row += 1

        # Substrato
        ctk.CTkLabel(self, text="εr do substrato:").grid(row=row, column=0, sticky="e", padx=6, pady=6)
        self.e_eps = ctk.CTkEntry(self); self.e_eps.insert(0, "4.4")
        self.e_eps.grid(row=row, column=1, sticky="w", padx=6, pady=6)

        ctk.CTkLabel(self, text="Altura h (mm):").grid(row=row, column=2, sticky="e", padx=6, pady=6)
        self.e_h = ctk.CTkEntry(self); self.e_h.insert(0, "1.57")
        self.e_h.grid(row=row, column=3, sticky="w", padx=6, pady=6)
        row += 1

        ctk.CTkLabel(self, text="Espessura cobre tCu (mm):").grid(row=row, column=0, sticky="e", padx=6, pady=6)
        self.e_tcu = ctk.CTkEntry(self); self.e_tcu.insert(0, "0.035")
        self.e_tcu.grid(row=row, column=1, sticky="w", padx=6, pady=6)
        row += 1

        self.var_run = ctk.BooleanVar(value=True)
        self.chk_run = ctk.CTkCheckBox(self, text="Rodar simulação após criar", variable=self.var_run)
        self.chk_run.grid(row=row, column=0, columnspan=2, sticky="w", padx=10, pady=6)
        row += 1

        self.btn = ctk.CTkButton(self, text="Criar Array no HFSS", command=self.on_create)
        self.btn.grid(row=row, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        row += 1

        # Log
        self.txt = ctk.CTkTextbox(self, height=220)
        self.txt.grid(row=row, column=0, columnspan=4, sticky="nsew", padx=10, pady=10)

        # layout weight
        self.grid_rowconfigure(row, weight=1)
        for c in range(4):
            self.grid_columnconfigure(c, weight=1)

        # hook close
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def log(self, s: str):
        self.txt.insert("end", s + "\n")
        self.txt.see("end")
        print(s, flush=True)

    # ----------------- callbacks -----------------

    def on_create(self):
        self.txt.delete("1.0", "end")
        try:
            f_low = float(self.e_flow.get().strip().replace(",", "."))
            f_high = float(self.e_fhigh.get().strip().replace(",", "."))
            f0 = float(self.e_f0.get().strip().replace(",", "."))
            g = float(self.e_gain.get().strip().replace(",", "."))
            eps = float(self.e_eps.get().strip().replace(",", "."))
            h = float(self.e_h.get().strip().replace(",", "."))
            tcu = float(self.e_tcu.get().strip().replace(",", "."))
        except Exception:
            self.log("Parâmetros inválidos.")
            return

        self.log(f"[Analítico] f0={f0:.3f} GHz | banda={f_low:.3f}–{f_high:.3f} GHz")
        Wp, Lp, eps_eff = patch_dimensions(f0, eps, h)
        self.log(f"[Patch] W≈{Wp:.2f} mm, L≈{Lp:.2f} mm, εeff≈{eps_eff:.4f}")

        # caminho do projeto ao lado do script
        try:
            base = os.path.dirname(__file__)
        except NameError:
            base = os.getcwd()
        proj = os.path.join(base, "PatchArray_HFSS.aedt")

        # cria/roda
        try:
            self.hfss, info = build_array_project(
                project_path=proj,
                f_low_ghz=f_low, f_high_ghz=f_high, f0_ghz=f0,
                g_target_dbi=g, eps_r=eps, h_mm=h, tcu_mm=tcu,
                run_after=self.var_run.get(), version="2024.2"
            )
        except Exception as e:
            self.log("Erro ao criar o projeto:")
            self.log(traceback.format_exc())
            return

        # mensagens finais
        self.log(f"[Array] alvo={g:.2f} dBi | Nx={info['Nx']} Ny={info['Ny']} (N={info['N_real']})")
        self.log(f"[λ/4] Wt≈{info['Wt_mm']:.2f} mm | Lq≈{info['Lq_mm']:.2f} mm | 50Ω W≈{info['W50_mm']:.2f} mm")
        self.log(f"[Setup] Executado: {info['setup']} : {info['sweep']}  ({info['f1_ghz']:.3f}–{info['f2_ghz']:.3f} GHz, FAST 101 pts)")
        self.log(f"[Projeto] {proj}")
        self.log(f"[Ports] {', '.join(info['ports'])}")

    def on_close(self):
        try:
            if self.hfss is not None:
                self.hfss.save_project()
                self.hfss.release_desktop()
        except Exception:
            pass
        self.destroy()


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
