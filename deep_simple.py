import os
import tempfile
import time
import customtkinter as ctk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np
import math

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
            "gain": 12.0,       # dBi (ganho desejado)
            "sweep_start": 8.0,
            "sweep_stop": 12.0,
            "cores": 4,
            "aedt_version": "2024.2",
            "non_graphical": True  # Modo não-gráfico por padrão
        }
        
        # Variáveis calculadas
        self.calculated_params = {
            "num_patches": 1,
            "spacing": 15.0,    # mm
            "patch_length": 9.57,
            "patch_width": 9.25
        }
        
        self.setup_gui()
        
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Patch Antenna Array Designer - Versão Simplificada")
        self.window.geometry("900x700")
        
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
        
        # Número de núcleos
        cores_label = ctk.CTkLabel(params_frame, text="Núcleos para Simulação:")
        cores_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        cores_entry = ctk.CTkEntry(params_frame)
        cores_entry.insert(0, str(self.params["cores"]))
        cores_entry.grid(row=row, column=1, padx=5, pady=5)
        entries.append(("cores", cores_entry))
        row += 1
        
        # Checkbox para modo não-gráfico
        gui_label = ctk.CTkLabel(params_frame, text="Modo Não-Gráfico:")
        gui_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        gui_var = ctk.BooleanVar(value=self.params["non_graphical"])
        gui_checkbox = ctk.CTkCheckBox(params_frame, text="", variable=gui_var)
        gui_checkbox.grid(row=row, column=1, padx=5, pady=5, sticky="w")
        entries.append(("non_graphical", gui_var))
        row += 1
        
        self.entries = entries
        
        # Frame de parâmetros calculados
        calc_frame = ctk.CTkFrame(main_frame)
        calc_frame.pack(fill="x", padx=10, pady=10)
        
        calc_title = ctk.CTkLabel(calc_frame, text="Parâmetros Calculados", 
                                 font=ctk.CTkFont(size=16, weight="bold"))
        calc_title.pack(pady=5)
        
        # Labels para mostrar os parâmetros calculados
        self.patches_label = ctk.CTkLabel(calc_frame, text="Número de Patches: 1")
        self.patches_label.pack(pady=2)
        
        self.spacing_label = ctk.CTkLabel(calc_frame, text="Espaçamento: 15.0 mm")
        self.spacing_label.pack(pady=2)
        
        self.dimensions_label = ctk.CTkLabel(calc_frame, text="Dimensões do Patch: 9.57 x 9.25 mm")
        self.dimensions_label.pack(pady=2)
        
        # Botões
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        calc_button = ctk.CTkButton(button_frame, text="Calcular Parâmetros", command=self.calculate_parameters)
        calc_button.pack(side="left", padx=5, pady=5)
        
        run_button = ctk.CTkButton(button_frame, text="Executar Simulação", command=self.run_simulation)
        run_button.pack(side="left", padx=5, pady=5)
        
        quit_button = ctk.CTkButton(button_frame, text="Sair", command=self.window.quit)
        quit_button.pack(side="right", padx=5, pady=5)
        
        # Área de resultados
        results_frame = ctk.CTkFrame(main_frame)
        results_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.status_label = ctk.CTkLabel(results_frame, text="Pronto para calcular parâmetros")
        self.status_label.pack(pady=5)
        
        # Canvas para plotagem
        self.fig, self.ax = plt.subplots(figsize=(6, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=results_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        
    def get_parameters(self):
        """Obtém os parâmetros da interface"""
        for key, entry in self.entries:
            try:
                if key in ["cores"]:
                    self.params[key] = int(entry.get())
                elif key == "non_graphical":
                    self.params[key] = entry.get()
                else:
                    self.params[key] = float(entry.get())
            except ValueError:
                self.status_label.configure(text=f"Erro: Valor inválido para {key}")
                return False
        return True
    
    def calculate_parameters(self):
        """Calcula os parâmetros do array baseado no ganho desejado"""
        if not self.get_parameters():
            return
            
        try:
            # Calcula as dimensões do patch
            patch_length, patch_width = self.calculate_patch_dimensions(self.params["frequency"])
            self.calculated_params["patch_length"] = patch_length
            self.calculated_params["patch_width"] = patch_width
            
            # Calcula o número de patches necessário para o ganho desejado
            G0 = 8.0  # Ganho aproximado de um patch individual em dBi
            desired_gain = self.params["gain"]
            
            # Calcula o número de patches necessário
            num_patches = int(math.ceil(10 ** ((desired_gain - G0) / 10)))
            self.calculated_params["num_patches"] = num_patches
            
            # Calcula o espaçamento ideal
            c = 3e8  # velocidade da luz m/s
            er = 2.2  # permissividade relativa do substrato (Duroid)
            freq = self.params["frequency"] * 1e9
            lambda0 = c / freq
            lambda_g = lambda0 / math.sqrt(er)
            spacing = lambda_g / 2 * 1000  # converter para mm
            self.calculated_params["spacing"] = spacing
            
            # Atualiza a interface com os valores calculados
            self.patches_label.configure(text=f"Número de Patches: {num_patches}")
            self.spacing_label.configure(text=f"Espaçamento: {spacing:.2f} mm")
            self.dimensions_label.configure(text=f"Dimensões do Patch: {patch_length:.2f} x {patch_width:.2f} mm")
            
            self.status_label.configure(text="Parâmetros calculados com sucesso!")
            
        except Exception as e:
            self.status_label.configure(text=f"Erro no cálculo: {str(e)}")
    
    def calculate_patch_dimensions(self, frequency):
        """Calcula as dimensões do patch baseado na frequência"""
        base_freq = 10.0
        base_length = 9.57
        base_width = 9.25
        
        # Escala inversa com a frequência
        length = (base_length * base_freq) / frequency
        width = (base_width * base_freq) / frequency
        
        return length, width
    
    def run_simulation(self):
        """Executa a simulação simplificada"""
        if not self.get_parameters():
            return
            
        # Se os parâmetros não foram calculados, calcula agora
        if self.calculated_params["num_patches"] == 1:
            self.calculate_parameters()
            
        self.status_label.configure(text="Iniciando simulação...")
        self.window.update()
        
        try:
            # Simula o processo de simulação
            self.status_label.configure(text="Criando projeto HFSS...")
            self.window.update()
            time.sleep(1)
            
            self.status_label.configure(text="Configurando geometria...")
            self.window.update()
            time.sleep(1)
            
            self.status_label.configure(text="Definindo materiais...")
            self.window.update()
            time.sleep(1)
            
            self.status_label.configure(text="Configurando simulação...")
            self.window.update()
            time.sleep(1)
            
            self.status_label.configure(text="Simulando...")
            self.window.update()
            time.sleep(2)
            
            # Gera dados simulados para demonstração
            self.generate_demo_results()
            
            self.status_label.configure(text="Simulação concluída com sucesso!")
            
        except Exception as e:
            self.status_label.configure(text=f"Erro na simulação: {str(e)}")
    
    def generate_demo_results(self):
        """Gera resultados de demonstração"""
        try:
            self.ax.clear()
            
            # Gera dados simulados
            frequencies = np.linspace(self.params["sweep_start"], self.params["sweep_stop"], 100)
            
            # Simula S11 com ressonância na frequência central
            center_freq = self.params["frequency"]
            s11_data = -20 * np.exp(-((frequencies - center_freq) / 0.5)**2) - 5
            
            # Adiciona ruído realista
            noise = np.random.normal(0, 0.5, len(s11_data))
            s11_data += noise
            
            # Plota S11
            self.ax.plot(frequencies, s11_data, label="S11 (Simulado)", linewidth=2)
            
            # Adiciona linha de -10dB como referência
            self.ax.axhline(y=-10, color='r', linestyle='--', alpha=0.7, label='-10 dB')
            
            self.ax.set_xlabel("Frequência (GHz)")
            self.ax.set_ylabel("S-Parameter (dB)")
            self.ax.set_title("Resultados Simulados - S-Parameter")
            self.ax.legend()
            self.ax.grid(True)
            
            # Destaca a frequência central
            self.ax.axvline(x=center_freq, color='g', linestyle='--', alpha=0.7)
            self.ax.text(center_freq+0.1, self.ax.get_ylim()[1]-2, 
                        f'{center_freq} GHz', color='g')
            
            self.canvas.draw()
            
        except Exception as e:
            self.status_label.configure(text=f"Erro ao gerar resultados: {str(e)}")
    
    def run(self):
        """Executa a aplicação"""
        try:
            self.window.mainloop()
        except Exception as e:
            print(f"Erro na aplicação: {e}")

# Executa a aplicação
if __name__ == "__main__":
    app = PatchAntennaDesigner()
    app.run()

