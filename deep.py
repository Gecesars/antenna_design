import os
import tempfile
import time
import ansys.aedt.core
from ansys.aedt.core.modeler.advanced_cad.stackup_3d import Stackup3D
import customtkinter as ctk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np

# Configuração da interface gráfica
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class PatchAntennaDesigner:
    def __init__(self):
        self.hfss = None
        self.temp_folder = None
        self.project_name = ""
        
        # Parâmetros padrão
        self.params = {
            "frequency": 10.0,  # GHz
            "gain": 8.0,        # dBi
            "sweep_start": 8.0,
            "sweep_stop": 12.0,
            "num_patches": 1,
            "spacing": 15.0,    # mm
            "cores": 4,
            "aedt_version": "2024.2"
        }
        
        self.setup_gui()
        
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer")
        self.window.geometry("800x600")
        
        # Frame principal
        main_frame = ctk.CTkFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Título
        title_label = ctk.CTkLabel(main_frame, text="Patch Antenna Array Designer", 
                                  font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=10)
        
        # Frame de parâmetros
        params_frame = ctk.CTkFrame(main_frame)
        params_frame.pack(fill="x", padx=10, pady=10)
        
        # Campos de entrada
        entries = []
        row = 0
        
        # Frequência central
        freq_label = ctk.CTkLabel(params_frame, text="Frequência Central (GHz):")
        freq_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        freq_entry = ctk.CTkEntry(params_frame)
        freq_entry.insert(0, str(self.params["frequency"]))
        freq_entry.grid(row=row, column=1, padx=5, pady=5)
        entries.append(("frequency", freq_entry))
        row += 1
        
        # Ganho desejado
        gain_label = ctk.CTkLabel(params_frame, text="Ganho Desejado (dBi):")
        gain_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        gain_entry = ctk.CTkEntry(params_frame)
        gain_entry.insert(0, str(self.params["gain"]))
        gain_entry.grid(row=row, column=1, padx=5, pady=5)
        entries.append(("gain", gain_entry))
        row += 1
        
        # Início do sweep
        sweep_start_label = ctk.CTkLabel(params_frame, text="Início do Sweep (GHz):")
        sweep_start_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        sweep_start_entry = ctk.CTkEntry(params_frame)
        sweep_start_entry.insert(0, str(self.params["sweep_start"]))
        sweep_start_entry.grid(row=row, column=1, padx=5, pady=5)
        entries.append(("sweep_start", sweep_start_entry))
        row += 1
        
        # Fim do sweep
        sweep_stop_label = ctk.CTkLabel(params_frame, text="Fim do Sweep (GHz):")
        sweep_stop_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        sweep_stop_entry = ctk.CTkEntry(params_frame)
        sweep_stop_entry.insert(0, str(self.params["sweep_stop"]))
        sweep_stop_entry.grid(row=row, column=1, padx=5, pady=5)
        entries.append(("sweep_stop", sweep_stop_entry))
        row += 1
        
        # Número de patches
        num_patches_label = ctk.CTkLabel(params_frame, text="Número de Patches:")
        num_patches_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        num_patches_entry = ctk.CTkEntry(params_frame)
        num_patches_entry.insert(0, str(self.params["num_patches"]))
        num_patches_entry.grid(row=row, column=1, padx=5, pady=5)
        entries.append(("num_patches", num_patches_entry))
        row += 1
        
        # Espaçamento entre patches
        spacing_label = ctk.CTkLabel(params_frame, text="Espaçamento (mm):")
        spacing_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        spacing_entry = ctk.CTkEntry(params_frame)
        spacing_entry.insert(0, str(self.params["spacing"]))
        spacing_entry.grid(row=row, column=1, padx=5, pady=5)
        entries.append(("spacing", spacing_entry))
        row += 1
        
        # Número de núcleos
        cores_label = ctk.CTkLabel(params_frame, text="Núcleos para Simulação:")
        cores_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        cores_entry = ctk.CTkEntry(params_frame)
        cores_entry.insert(0, str(self.params["cores"]))
        cores_entry.grid(row=row, column=1, padx=5, pady=5)
        entries.append(("cores", cores_entry))
        row += 1
        
        self.entries = entries
        
        # Botões
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        run_button = ctk.CTkButton(button_frame, text="Executar Simulação", command=self.run_simulation)
        run_button.pack(side="left", padx=5, pady=5)
        
        quit_button = ctk.CTkButton(button_frame, text="Sair", command=self.window.quit)
        quit_button.pack(side="right", padx=5, pady=5)
        
        # Área de resultados
        results_frame = ctk.CTkFrame(main_frame)
        results_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.status_label = ctk.CTkLabel(results_frame, text="Pronto para simular")
        self.status_label.pack(pady=5)
        
        # Canvas para plotagem
        self.fig, self.ax = plt.subplots(figsize=(6, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=results_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        
    def get_parameters(self):
        """Obtém os parâmetros da interface"""
        for key, entry in self.entries:
            try:
                if key in ["num_patches", "cores"]:
                    self.params[key] = int(entry.get())
                else:
                    self.params[key] = float(entry.get())
            except ValueError:
                self.status_label.configure(text=f"Erro: Valor inválido para {key}")
                return False
        return True
    
    def calculate_patch_dimensions(self, frequency):
        """Calcula as dimensões do patch baseado na frequência"""
        # Fórmulas simplificadas para patch antenna
        c = 3e8  # velocidade da luz m/s
        er = 2.2  # permissividade relativa do substrato (Duroid)
        
        # Comprimento efetivo
        lambda0 = c / (frequency * 1e9)
        lambda_g = lambda0 / np.sqrt(er)
        
        # Dimensões do patch (aproximadas)
        length = lambda_g / 2 * 1000  # convert to mm
        width = length * 0.9  # relação aspect ratio comum
        
        return length, width
    
    def run_simulation(self):
        """Executa a simulação completa"""
        if not self.get_parameters():
            return
            
        self.status_label.configure(text="Iniciando simulação...")
        self.window.update()
        
        try:
            # Cria diretório temporário
            self.temp_folder = tempfile.TemporaryDirectory(suffix=".ansys")
            self.project_name = os.path.join(self.temp_folder.name, "patch_array.aedt")
            
            # Inicializa HFSS
            self.hfss = ansys.aedt.core.Hfss(
                project=self.project_name,
                solution_type="Terminal",
                design="patch_array",
                non_graphical=True,
                new_desktop=True,
                version=self.params["aedt_version"],
            )
            
            # Configura unidades
            length_units = "mm"
            self.hfss.modeler.model_units = length_units
            
            # Cria stackup
            stackup = Stackup3D(self.hfss)
            ground = stackup.add_ground_layer(
                "ground", material="copper", thickness=0.035, fill_material="air"
            )
            dielectric = stackup.add_dielectric_layer(
                "dielectric", thickness="0.5" + length_units, material="Duroid (tm)"
            )
            signal = stackup.add_signal_layer(
                "signal", material="copper", thickness=0.035, fill_material="air"
            )
            
            # Calcula dimensões do patch
            patch_length, patch_width = self.calculate_patch_dimensions(self.params["frequency"])
            
            # Cria múltiplos patches
            patches = []
            spacing = self.params["spacing"]
            num_patches = self.params["num_patches"]
            
            for i in range(num_patches):
                x_offset = i * (patch_width + spacing)
                patch = signal.add_patch(
                    patch_length=patch_length, 
                    patch_width=patch_width, 
                    patch_name=f"Patch_{i+1}",
                    frequency=self.params["frequency"] * 1e9,
                    origin=[x_offset, 0, 0]
                )
                patch.create_probe_port(ground, rel_x_offset=0.485)
                patches.append(patch)
            
            # Redimensiona a região em torno do array
            stackup.resize_around_element(patches[0])
            pad_length = [3, 3, 3, 3, 3, 3]  # Air bounding box buffer in mm.
            region = self.hfss.modeler.create_region(pad_length, is_percentage=False)
            self.hfss.assign_radiation_boundary_to_objects(region)
            
            # Define configuração de simulação
            setup = self.hfss.create_setup(
                name="Setup1", 
                setup_type="HFSSDriven", 
                Frequency=f"{self.params['frequency']}GHz"
            )
            
            setup.create_frequency_sweep(
                unit="GHz",
                name="Sweep1",
                start_frequency=self.params["sweep_start"],
                stop_frequency=self.params["sweep_stop"],
                sweep_type="Interpolating",
            )
            
            self.status_label.configure(text="Simulando...")
            self.window.update()
            
            # Executa a simulação
            self.hfss.save_project()
            self.hfss.analyze(cores=self.params["cores"])
            
            # Processa resultados
            self.status_label.configure(text="Processando resultados...")
            self.window.update()
            
            self.plot_results()
            
            self.status_label.configure(text="Simulação concluída com sucesso!")
            
        except Exception as e:
            self.status_label.configure(text=f"Erro: {str(e)}")
            
    def plot_results(self):
        """Plota os resultados da simulação"""
        self.ax.clear()
        
        # Obtém dados S-parameter
        plot_data = self.hfss.get_traces_for_plot()
        report = self.hfss.post.create_report(plot_data)
        solution = report.get_solution_data()
        
        # Plota S11 para cada porta
        for i, trace in enumerate(solution.expressions):
            if "S" in trace and "dB" in trace:
                freq_data = solution.sweeps["Freq"]
                s_data = solution.data_real(trace)
                self.ax.plot(freq_data, s_data, label=f"Porta {i+1}")
        
        self.ax.set_xlabel("Frequência (GHz)")
        self.ax.set_ylabel("S-Parameter (dB)")
        self.ax.set_title("Resultados de S-Parameter")
        self.ax.legend()
        self.ax.grid(True)
        
        self.canvas.draw()
        
    def cleanup(self):
        """Limpa recursos após fechar a aplicação"""
        if self.hfss:
            try:
                self.hfss.save_project()
                self.hfss.release_desktop()
                time.sleep(3)
            except:
                pass
                
        if self.temp_folder:
            try:
                self.temp_folder.cleanup()
            except:
                pass
    
    def run(self):
        """Executa a aplicação"""
        try:
            self.window.mainloop()
        finally:
            self.cleanup()

# Executa a aplicação
if __name__ == "__main__":
    app = PatchAntennaDesigner()
    app.run()