#!/usr/bin/env python3
"""
Script de teste para verificar a instalação do PyAEDT
"""

import sys
import os

def test_imports():
    """Testa se os módulos necessários podem ser importados"""
    print("=== Teste de Imports ===")
    
    try:
        import ansys.aedt.core
        print("✓ ansys.aedt.core importado com sucesso")
        print(f"  Versão: {ansys.aedt.core.__version__}")
    except ImportError as e:
        print(f"✗ Erro ao importar ansys.aedt.core: {e}")
        return False
    
    try:
        from ansys.aedt.core.modeler.advanced_cad.stackup_3d import Stackup3D
        print("✓ Stackup3D importado com sucesso")
    except ImportError as e:
        print(f"✗ Erro ao importar Stackup3D: {e}")
        return False
    
    try:
        import customtkinter as ctk
        print("✓ customtkinter importado com sucesso")
    except ImportError as e:
        print(f"✗ Erro ao importar customtkinter: {e}")
        return False
    
    try:
        import matplotlib.pyplot as plt
        print("✓ matplotlib importado com sucesso")
    except ImportError as e:
        print(f"✗ Erro ao importar matplotlib: {e}")
        return False
    
    try:
        import numpy as np
        print("✓ numpy importado com sucesso")
    except ImportError as e:
        print(f"✗ Erro ao importar numpy: {e}")
        return False
    
    return True

def test_hfss_connection():
    """Testa a conexão com o HFSS"""
    print("\n=== Teste de Conexão HFSS ===")
    
    try:
        import ansys.aedt.core
        
        # Tenta conectar ao HFSS
        print("Tentando conectar ao HFSS...")
        
        # Primeiro, vamos verificar se o AEDT está disponível
        try:
            hfss = ansys.aedt.core.Hfss(
                project="test_project",
                solution_type="Terminal",
                design="test_design",
                non_graphical=True,
                new_desktop=True,
                version="2024.2"
            )
            print("✓ HFSS inicializado com sucesso")
            
            # Fecha o projeto de teste
            hfss.release_desktop()
            print("✓ Projeto de teste fechado")
            return True
            
        except Exception as e:
            print(f"✗ Erro ao inicializar HFSS: {e}")
            print("  Isso pode indicar:")
            print("  - ANSYS AEDT não está instalado")
            print("  - Versão incorreta do AEDT")
            print("  - Problemas de licença")
            return False
            
    except Exception as e:
        print(f"✗ Erro geral: {e}")
        return False

def test_simple_gui():
    """Testa a interface gráfica básica"""
    print("\n=== Teste de Interface Gráfica ===")
    
    try:
        import customtkinter as ctk
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        # Cria uma janela simples
        root = ctk.CTk()
        root.title("Teste de Interface")
        root.geometry("400x300")
        
        # Adiciona um label
        label = ctk.CTkLabel(root, text="Interface funcionando!")
        label.pack(pady=20)
        
        # Cria um gráfico simples
        fig, ax = plt.subplots(figsize=(4, 3))
        x = [1, 2, 3, 4, 5]
        y = [1, 4, 2, 5, 3]
        ax.plot(x, y)
        ax.set_title("Gráfico de Teste")
        
        canvas = FigureCanvasTkAgg(fig, master=root)
        canvas.get_tk_widget().pack()
        
        print("✓ Interface gráfica criada com sucesso")
        print("  (Janela será fechada automaticamente em 3 segundos)")
        
        # Fecha a janela após 3 segundos
        root.after(3000, root.destroy)
        root.mainloop()
        
        return True
        
    except Exception as e:
        print(f"✗ Erro na interface gráfica: {e}")
        return False

def main():
    """Função principal"""
    print("Teste de Instalação do PyAEDT")
    print("=" * 40)
    
    # Testa imports
    if not test_imports():
        print("\n❌ Falha nos imports. Verifique a instalação.")
        return False
    
    # Testa interface gráfica
    if not test_simple_gui():
        print("\n❌ Falha na interface gráfica.")
        return False
    
    # Testa conexão HFSS (opcional)
    print("\nDeseja testar a conexão com o HFSS? (s/n): ", end="")
    try:
        response = input().lower()
        if response in ['s', 'sim', 'y', 'yes']:
            test_hfss_connection()
    except:
        pass
    
    print("\n✅ Testes básicos concluídos!")
    print("\nPróximos passos:")
    print("1. Execute 'python deep_simple.py' para testar a versão simplificada")
    print("2. Se funcionar, tente 'python deep.py' para a versão completa")
    
    return True

if __name__ == "__main__":
    main()

