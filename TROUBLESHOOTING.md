# Guia de Solução de Problemas - Patch Antenna Designer

## Análise dos Erros Encontrados

### 1. Erro Principal: `'nonetype' object has no attribute 'getname'`

**Causa:** Problema na inicialização do HFSS com PyAEDT versão 0.18.1

**Soluções:**

#### Solução 1: Atualizar PyAEDT
```bash
pip install --upgrade pyaedt
```

#### Solução 2: Usar versão específica compatível
```bash
pip install pyaedt==0.17.0
```

#### Solução 3: Verificar instalação do ANSYS AEDT
- Certifique-se de que o ANSYS Electronics Desktop 2024.2 está instalado
- Verifique se a licença está ativa
- Confirme que o caminho está nas variáveis de ambiente

### 2. Erro: `'bool' object has no attribute 'getname'`

**Causa:** Parâmetro `non_graphical` sendo passado incorretamente

**Solução:** Já corrigido no código atualizado

### 3. Erros de Tkinter

**Causa:** Conflitos entre CustomTkinter e interface do HFSS

**Soluções:**

#### Solução 1: Usar modo não-gráfico
```python
non_graphical=True
```

#### Solução 2: Executar em ambiente separado
```bash
# Em um terminal separado
python deep_simple.py
```

## Arquivos de Teste

### 1. `test_pyaedt.py`
Script para verificar a instalação:
```bash
python test_pyaedt.py
```

### 2. `deep_simple.py`
Versão simplificada sem HFSS:
```bash
python deep_simple.py
```

### 3. `deep.py`
Versão completa com HFSS (após resolver problemas):
```bash
python deep.py
```

## Passos para Resolver

### Passo 1: Teste Básico
```bash
python test_pyaedt.py
```

### Passo 2: Teste Interface
```bash
python deep_simple.py
```

### Passo 3: Verificar ANSYS
1. Abra o ANSYS Electronics Desktop manualmente
2. Verifique se funciona normalmente
3. Confirme a versão instalada

### Passo 4: Teste HFSS
```bash
python test_pyaedt.py
# Responda 's' quando perguntado sobre testar HFSS
```

### Passo 5: Executar Versão Completa
```bash
python deep.py
```

## Configurações Recomendadas

### Ambiente Python
```bash
# Criar ambiente virtual
python -m venv venv

# Ativar ambiente
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Instalar dependências
pip install pyaedt customtkinter matplotlib numpy
```

### Versões Testadas
- Python: 3.12.4
- PyAEDT: 0.18.1
- CustomTkinter: 5.2.0
- Matplotlib: 3.8.0
- NumPy: 1.24.0

## Problemas Comuns

### 1. ANSYS não encontrado
**Sintoma:** `AEDT installation Path not found`

**Solução:**
- Verificar instalação do ANSYS
- Adicionar ao PATH do sistema
- Usar caminho completo na inicialização

### 2. Erro de licença
**Sintoma:** `License error` ou `Failed to start AEDT`

**Solução:**
- Verificar licença ANSYS
- Reiniciar servidor de licenças
- Contatar administrador de licenças

### 3. Conflito de versões
**Sintoma:** `Version mismatch` ou `Incompatible version`

**Solução:**
- Usar versão compatível do PyAEDT
- Verificar versão do ANSYS AEDT
- Atualizar ou fazer downgrade conforme necessário

## Logs e Debug

### Habilitar logs detalhados
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

### Verificar logs do PyAEDT
```python
import ansys.aedt.core
ansys.aedt.core.settings.enable_desktop_logs = True
```

### Logs do sistema
- Windows: `%TEMP%\pyaedt_*.log`
- Linux: `/tmp/pyaedt_*.log`

## Alternativas

### 1. Usar versão simplificada
O arquivo `deep_simple.py` funciona sem HFSS e pode ser usado para:
- Testar cálculos
- Verificar interface
- Demonstrar funcionalidades

### 2. Usar ANSYS diretamente
- Abrir ANSYS Electronics Desktop
- Criar projeto manualmente
- Usar scripts Python separados

### 3. Usar outras ferramentas
- CST Studio Suite
- COMSOL Multiphysics
- FEKO

## Suporte

### Recursos úteis
- [Documentação PyAEDT](https://aedt.docs.pyansys.com/)
- [Fórum ANSYS](https://forum.ansys.com/)
- [GitHub PyAEDT](https://github.com/ansys/pyaedt)

### Comandos de diagnóstico
```bash
# Verificar versões
pip list | grep -E "(pyaedt|ansys)"

# Verificar instalação ANSYS
where ansysedt.exe

# Testar conexão
python -c "import ansys.aedt.core; print('OK')"
```

## Conclusão

Os erros encontrados são principalmente relacionados à:
1. Incompatibilidade de versões
2. Configuração incorreta de parâmetros
3. Problemas de instalação do ANSYS

A versão simplificada (`deep_simple.py`) deve funcionar imediatamente e pode ser usada para testar a funcionalidade básica enquanto os problemas do HFSS são resolvidos.

