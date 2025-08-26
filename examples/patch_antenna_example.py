import os
import sys
import math
import shutil
import logging

# PyAEDT
from ansys.aedt.core import Hfss

# ----------------------------- Configurações ---------------------------------
AEDT_VERSION = "2024.2"
FREQ_GHZ = 2.4
SUBSTRATE_HEIGHT_MM = 1.57
EPS_R = 4.4
COPPER_THICK_MM = 0.035

PIN_RADIUS_MM = 0.6
PIN_GAP_MM = 0.5
FEED_OFFSET_X_MM = 8.0

SWEEP_START_GHZ = 1.5
SWEEP_STOP_GHZ = 3.5
SWEEP_POINTS = 101
SWEEP_NAME = "FastSweep"
SETUP_NAME = "MainSetup"
DESIGN_NAME = "PatchAntenna_DirectHFSS"
S_PARAM_EXPR = "db(S(1,1))"

# ---------------------- Diretórios e Logging ---------------------------
try:
    OUTPUT_DIR = os.path.dirname(__file__)
except NameError:
    OUTPUT_DIR = os.getcwd()

PROJECT_PATH = os.path.join(OUTPUT_DIR, f"{DESIGN_NAME}.aedt")
LOG_PATH = os.path.join(OUTPUT_DIR, "run.log")
CSV_PATH = os.path.join(OUTPUT_DIR, "s11_results.csv")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, mode='w', encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("antenna_suite")

# ---------------------- Fórmulas e Utilidades ---------------------
def design_patch_dimensions(f0_GHz: float, eps_r: float, h_mm: float):
    c = 299792458.0; f0 = f0_GHz * 1e9; h = h_mm / 1000.0
    W = (c / (2.0 * f0)) * math.sqrt(2.0 / (eps_r + 1.0))
    eps_eff = ((eps_r + 1.0) / 2.0) + ((eps_r - 1.0) / 2.0) * (1.0 + 12.0 * h / W) ** -0.5
    dL_over_h = 0.412 * ((eps_eff + 0.3) * (W / h + 0.264)) / ((eps_eff - 0.258) * (W / h + 0.8))
    dL = dL_over_h * h; L_eff = c / (2.0 * f0 * math.sqrt(eps_eff)); L = L_eff - 2.0 * dL
    return W * 1000.0, L * 1000.0

def clean_previous_project(project_path: str):
    if os.path.exists(project_path):
        log.info("Removendo projeto antigo...")
        try: os.remove(project_path)
        except Exception as e: log.warning(f"Não foi possível remover {project_path}: {e}")
    lock_file = project_path + ".lock"
    if os.path.exists(lock_file):
        try: os.remove(lock_file)
        except Exception as e: log.warning(f"Não foi possível remover {lock_file}: {e}")
    results_folder = project_path.replace(".aedt", ".aedtresults")
    if os.path.exists(results_folder):
        try: shutil.rmtree(results_folder)
        except Exception as e: log.warning(f"Não foi possível remover {results_folder}: {e}")

# ------------------------------- Fluxo principal -----------------------------
def main():
    log.info("========== Antenna Automation (Direct HFSS Modeling) ==========")
    PATCH_WIDTH_MM, PATCH_LENGTH_MM = design_patch_dimensions(FREQ_GHZ, EPS_R, SUBSTRATE_HEIGHT_MM)
    log.info(f"[Analítico] W≈{PATCH_WIDTH_MM:.2f} mm, L≈{PATCH_LENGTH_MM:.2f} mm")

    clean_previous_project(PROJECT_PATH)

    hfss = Hfss(
        project=PROJECT_PATH, solution_type="Terminal", design=DESIGN_NAME,
        non_graphical=False, new_desktop=True, version=AEDT_VERSION,
    )
    
    try:
        hfss.modeler.model_units = "mm"
        
        log.info("Criando geometria...")
        substrate = hfss.modeler.create_box(
            origin=[f"{-PATCH_WIDTH_MM*1.5}", f"{-PATCH_LENGTH_MM*1.5}", "0"],
            sizes=[f"{PATCH_WIDTH_MM*3}", f"{PATCH_LENGTH_MM*3}", f"{-SUBSTRATE_HEIGHT_MM}"],
            name="Substrate", material="FR4_epoxy"
        )
        gnd = hfss.modeler.create_rectangle(
            origin=[f"{-PATCH_WIDTH_MM*1.5}", f"{-PATCH_LENGTH_MM*1.5}", f"{-SUBSTRATE_HEIGHT_MM}"],
            sizes=[f"{PATCH_WIDTH_MM*3}", f"{PATCH_LENGTH_MM*3}"],
            name="Ground", orientation="Z"
        )
        patch = hfss.modeler.create_box(
            origin=[f"{-PATCH_WIDTH_MM/2}", f"{-PATCH_LENGTH_MM/2}", "0"],
            sizes=[f"{PATCH_WIDTH_MM}", f"{PATCH_LENGTH_MM}", f"{COPPER_THICK_MM}"],
            name="Patch", material="copper"
        )

        log.info("Criando alimentação (probe feed)...")
        gnd_hole = hfss.modeler.create_circle(
            origin=[f"{FEED_OFFSET_X_MM}", "0", f"{-SUBSTRATE_HEIGHT_MM}"],
            radius=PIN_RADIUS_MM + PIN_GAP_MM,
            name="GND_Hole", orientation="Z"
        )
        hfss.modeler.subtract(gnd, gnd_hole, keep_originals=False)
        
        pin = hfss.modeler.create_cylinder(
            origin=[f"{FEED_OFFSET_X_MM}", "0", f"{-SUBSTRATE_HEIGHT_MM}"],
            radius=PIN_RADIUS_MM,
            height=f"{SUBSTRATE_HEIGHT_MM}",
            name="Pin", material="copper",
            orientation="Z"
        )
        
        port_cap = hfss.modeler.create_circle(
            origin=[f"{FEED_OFFSET_X_MM}", "0", f"{-SUBSTRATE_HEIGHT_MM}"],
            radius=PIN_RADIUS_MM + PIN_GAP_MM,
            name="Port_Cap", orientation="Z"
        )
        
        hfss.modeler.unite([patch, pin])
        
        log.info("Atribuindo contornos e excitação...")
        hfss.assign_perfecte_to_sheets([gnd.name, patch.name])
        
        # CORREÇÃO FINAL: Nome do método alterado para 'wave_port'.
        hfss.wave_port(
            sheet=port_cap,
            port_name="1",
            integration_line_start=[f"{FEED_OFFSET_X_MM}", f"{-(PIN_RADIUS_MM)}", f"{-SUBSTRATE_HEIGHT_MM}"],
            integration_line_stop=[f"{FEED_OFFSET_X_MM}", f"{(PIN_RADIUS_MM)}", f"{-SUBSTRATE_HEIGHT_MM}"]
        )
        
        region = hfss.modeler.create_region(pad_percent=300)
        hfss.assign_radiation_boundary_to_objects(region)

        log.info("Configurando a análise...")
        setup = hfss.create_setup(name=SETUP_NAME, Frequency=f"{FREQ_GHZ}GHz")
        setup.create_frequency_sweep(
            unit="GHz", name=SWEEP_NAME,
            start_frequency=SWEEP_START_GHZ, stop_frequency=SWEEP_STOP_GHZ,
            sweep_type="Interpolating", num_of_freq_points=SWEEP_POINTS,
        )

        log.info("Iniciando simulação...")
        hfss.analyze(setup.name)
        log.info("Simulação finalizada.")
        
        log.info("Exportando dados...")
        solution_data = hfss.post.get_solution_data(
            expressions=S_PARAM_EXPR,
            setup_sweep_name=f"{SETUP_NAME} : {SWEEP_NAME}",
        )
        if solution_data:
            solution_data.export_data_to_csv(CSV_PATH)
            log.info(f"Resultados S11 exportados para: {CSV_PATH}")
        
        hfss.post.create_report(S_PARAM_EXPR)
        
        hfss.save_project()
        log.info(f"Projeto salvo em: {PROJECT_PATH}")

    finally:
        hfss.release_desktop()
        log.info("AEDT liberado. Execução concluída.")

if __name__ == "__main__":
    main()