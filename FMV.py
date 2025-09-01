import customtkinter as ctk
import threading
import queue
import math
import os
import tempfile

# Tenta importar o PyAEDT. Se falhar, exibe uma mensagem de erro clara.
try:
    from pyaedt import Hfss
except ImportError:
    print("Erro: A biblioteca PyAEDT não foi encontrada.")
    print("Por favor, instale-a usando o comando: pip install pyaedt")
    exit()

class HfssFmvSimulatorApp(ctk.CTk):
    """
    Classe principal da aplicação com interface gráfica para simular uma antena FMV no HFSS.
    """

    def __init__(self):
        super().__init__()

        # --- Configurações da Janela Principal ---
        self.title("Gerador de Antena FMV para HFSS")
        self.geometry("600x650")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # --- Variáveis de Controle ---
        self.simulation_thread = None
        self.message_queue = queue.Queue()

        # --- Criação dos Widgets da GUI ---
        self._create_widgets()

        # Inicia o loop para processar mensagens da thread de simulação
        self.after(100, self._process_message_queue)

    def _create_widgets(self):
        """Cria e posiciona todos os widgets na janela."""
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # --- Seção de Entrada de Dados ---
        input_label = ctk.CTkLabel(main_frame, text="Configuração da Antena FMV", font=ctk.CTkFont(size=16, weight="bold"))
        input_label.pack(pady=(0, 15))

        freq_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        freq_frame.pack(pady=5, padx=10, fill="x")
        freq_label = ctk.CTkLabel(freq_frame, text="Frequência de Operação (MHz):", width=200, anchor="w")
        freq_label.pack(side="left")
        self.frequency_entry = ctk.CTkEntry(freq_frame, placeholder_text="Ex: 99.9")
        self.frequency_entry.pack(side="left", fill="x", expand=True)
        # Valor padrão comum para FM
        self.frequency_entry.insert(0, "99.9")

        # --- Botão de Ação ---
        self.run_button = ctk.CTkButton(main_frame, text="Gerar e Simular Antena no HFSS", command=self._start_simulation_thread)
        self.run_button.pack(pady=20, padx=10, ipady=5, fill="x")

        # --- Seção de Log de Saída ---
        log_label = ctk.CTkLabel(main_frame, text="Log da Simulação", font=ctk.CTkFont(size=14, weight="bold"))
        log_label.pack(pady=(10, 5))

        self.log_textbox = ctk.CTkTextbox(main_frame, state="disabled", height=400)
        self.log_textbox.pack(pady=10, padx=10, fill="both", expand=True)
        
        self._log_message("Bem-vindo! Insira a frequência e clique no botão para começar.")
        self._log_message("Baseado no 'Relatório Técnico Exaustivo: A Antena FMV'.")


    def _log_message(self, message):
        """Adiciona uma mensagem à caixa de texto de log de forma segura."""
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end") # Auto-scroll

    def _process_message_queue(self):
        """Verifica a fila de mensagens da thread e atualiza a GUI."""
        try:
            while True:
                message = self.message_queue.get_nowait()
                self._log_message(message)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._process_message_queue)

    def _start_simulation_thread(self):
        """Inicia a thread de automação do HFSS para não congelar a GUI."""
        frequency_str = self.frequency_entry.get()
        try:
            frequency_mhz = float(frequency_str)
            if not (87.5 <= frequency_mhz <= 108.1):
                self._log_message("ERRO: Frequência fora da faixa de FM comercial (87.5-108.1 MHz).")
                return
        except ValueError:
            self._log_message("ERRO: Frequência inválida. Por favor, insira um número.")
            return

        # Desabilita o botão para evitar múltiplas execuções
        self.run_button.configure(state="disabled", text="Simulando...")
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

        # Inicia a thread que fará o trabalho pesado
        self.simulation_thread = threading.Thread(
            target=self._hfss_worker,
            args=(frequency_mhz, self.message_queue, self.run_button)
        )
        self.simulation_thread.daemon = True
        self.simulation_thread.start()

    def _hfss_worker(self, frequency_mhz, msg_queue, button):
        """
        Função executada em uma thread separada para controlar o HFSS.
        Contém toda a lógica do PyAEDT. (VERSÃO CORRIGIDA)
        """
        try:
            # --- 1. Parâmetros da Antena (Baseado no Relatório Técnico) ---
            msg_queue.put(f"INFO: Frequência alvo: {frequency_mhz} MHz")

            # Fórmula prática do relatório para o comprimento de cada braço
            # L_braco(metros) = 71.5 / f(MHz)
            arm_length_m = 71.5 / frequency_mhz
            msg_queue.put(f"INFO: Comprimento calculado por braço: {arm_length_m:.4f} m")

            # Parâmetros Geométricos Críticos
            # O ângulo ótimo para impedância de 50 Ohms é frequentemente 90 graus.
            apex_angle_deg = 90.0
            wire_radius_mm = 1.0 # Raio do fio condutor
            feed_gap_mm = 1.0     # Pequeno espaço no centro para a porta de alimentação

            msg_queue.put("INFO: Parâmetros geométricos calculados com sucesso.")
            
            # --- 2. Automação do HFSS com PyAEDT ---
            project_name = f"FMV_Antenna_{frequency_mhz}MHz"
            design_name = "FMV_Analysis"
            
            # Diretório temporário para salvar o projeto
            temp_dir = tempfile.gettempdir()
            project_path = os.path.join(temp_dir, f"{project_name}.aedt")

            msg_queue.put("INFO: Iniciando e conectando-se ao Ansys Electronics Desktop...")
            # O 'with' garante que o AEDT seja fechado corretamente no final ou em caso de erro.
            # non_graphical=False permite ver a GUI do AEDT sendo construída.
            with Hfss(projectname=project_path, designname=design_name, specified_version="2024.2", non_graphical=False, new_desktop_session=True) as hfss:
                msg_queue.put("INFO: Conexão com HFSS estabelecida.")
                
                # --- 3. Criação da Geometria ---
                msg_queue.put("INFO: Criando a geometria 3D da antena...")
                
                # Criando os dois braços da antena como cilindros
                arm1 = hfss.modeler.create_cylinder(
                    cs_axis="Z",
                    position=[0, 0, feed_gap_mm / 2],
                    radius=wire_radius_mm,
                    height= -arm_length_m * 1000,
                    name="arm_1_temp",
                    matname="pec" # Perfect Electric Conductor
                )
                hfss.modeler.rotate(arm1, cs_axis="Y", angle=-(90-apex_angle_deg/2))

                arm2 = hfss.modeler.create_cylinder(
                    cs_axis="Z",
                    position=[0, 0, -feed_gap_mm / 2],
                    radius=wire_radius_mm,
                    height= -arm_length_m * 1000,
                    name="arm_2_temp",
                    matname="pec"
                )
                hfss.modeler.rotate(arm2, cs_axis="Y", angle=(90-apex_angle_deg/2))

                # CORREÇÃO 1: Unir os objetos e depois renomear.
                # O método unite usa os nomes dos objetos como uma lista de strings.
                hfss.modeler.unite([arm1.name, arm2.name])
                # O primeiro objeto da lista (arm1) agora é o objeto unido. Renomeamos ele.
                arm1.name = "Antenna_V"
                msg_queue.put("INFO: Geometria dos braços criada e unida.")

                # --- 4. Configuração da Excitação (Porta) ---
                msg_queue.put("INFO: Criando porta de alimentação (Lumped Port)...")
                # CORREÇÃO 2: A orientação do plano ("YZ") é o primeiro argumento POSICIONAL.
                port_sheet = hfss.modeler.create_rectangle(
                    "YZ", # Este é o argumento posicional 'orientation' que estava faltando
                    position=[0, -wire_radius_mm, -feed_gap_mm/2],
                    dimension_list=[wire_radius_mm*2, feed_gap_mm],
                    name="port_sheet"
                )
                hfss.lumped_port(port_sheet.name, impedance=50) # Impedância padrão de 50 Ohms
                msg_queue.put("INFO: Porta de alimentação configurada com 50 Ohms.")

                # --- 5. Configuração da Fronteira de Radiação ---
                msg_queue.put("INFO: Criando fronteira de radiação (arredores)...")
                # PyAEDT simplifica isso com uma função auxiliar
                hfss.create_open_region(Frequency=f"{frequency_mhz}MHz")
                msg_queue.put("INFO: Fronteira de radiação criada com sucesso.")

                # --- 6. Configuração da Análise ---
                msg_queue.put("INFO: Configurando a análise (Setup e Sweep)...")
                setup_name = "MainSetup"
                sweep_name = "FrequencySweep"
                
                setup = hfss.create_setup(setupname=setup_name)
                setup.props["Frequency"] = f"{frequency_mhz}MHz"
                setup.props["MaximumPasses"] = 10
                
                # Varredura de frequência para avaliar o VSWR
                hfss.create_linear_count_sweep(
                    setupname=setup_name,
                    unit="MHz",
                    freqstart=frequency_mhz * 0.8,
                    freqstop=frequency_mhz * 1.2,
                    num_of_freq_points=201,
                    sweepname=sweep_name,
                    sweep_type="Interpolating"
                )
                msg_queue.put("INFO: Análise configurada.")

                # --- 7. Execução da Simulação ---
                msg_queue.put("\n" + "="*40)
                msg_queue.put(f"INICIANDO SIMULAÇÃO: '{setup_name}'. Isso pode levar alguns minutos...")
                msg_queue.put("="*40 + "\n")
                hfss.analyze_setup(setup_name)
                msg_queue.put("SUCESSO: Simulação concluída.")
                
                # --- 8. Pós-processamento ---
                msg_queue.put("INFO: Gerando relatórios de resultados...")
                # Relatório do parâmetro S11 (relacionado ao VSWR)
                hfss.post.create_report(
                    "db(S(1,1))",
                    setup_name=f"{setup_name} : {sweep_name}",
                    primary_sweep_variable="Freq"
                )
                # Diagrama de irradiação 3D
                hfss.post.create_fieldplot_3d(
                    "GainTotal", 
                    setup_name, 
                    quantity_name="Gain", 
                    plot_name="3D_Gain_Pattern"
                )
                msg_queue.put("INFO: Relatórios de S11 e Ganho 3D criados no projeto HFSS.")
                
                # Salvar o projeto
                hfss.save_project()
                msg_queue.put(f"SUCESSO: Projeto salvo em: {project_path}")

        except Exception as e:
            msg_queue.put(f"\nERRO CRÍTICO: Ocorreu uma exceção durante a automação.")
            msg_queue.put(f"Detalhes: {str(e)}")
        finally:
            # Reabilita o botão na GUI, independentemente de sucesso ou falha
            button.configure(state="normal", text="Gerar e Simular Antena no HFSS")
            msg_queue.put("\nProcesso finalizado.")


if __name__ == "__main__":
    app = HfssFmvSimulatorApp()
    app.mainloop()