import pandas as pd
from tkinter import filedialog, messagebox
import threading

# Colunas de entrada
COLUNAS_ENTRADA = [
    "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
    "X", "Y", "X Corrigido", "Y Corrigido", "DAP", "Volumes (m³)", "X Negativo", "Y Negativo",
    "Latitude", "Longitude", "DM", "Observações", "N"
]

# Variável para armazenar a planilha principal
_planilha_principal = None  # Prefixo para indicar variável privada


def get_planilha_principal():
    """Retorna a planilha principal."""
    return _planilha_principal


def set_planilha_principal(planilha):
    """Define a planilha principal."""
    global _planilha_principal
    _planilha_principal = planilha


def selecionar_arquivos(entrada1_var, entrada2_var, carregar_planilha_principal):
    """Seleciona os arquivos das duas planilhas."""
    arquivo1 = filedialog.askopenfilename(
        title="Selecione a planilha de entrada (dados principais)",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if arquivo1:
        entrada1_var.set(arquivo1)
        # Carregar a planilha de forma assíncrona
        threading.Thread(target=carregar_planilha_principal, args=(arquivo1,)).start()

    arquivo2 = filedialog.askopenfilename(
        title="Selecione a segunda planilha (Nomes Vulgares e Científicos)",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if arquivo2:
        entrada2_var.set(arquivo2)


def carregar_planilha_principal(arquivo1):
    """Carrega a planilha principal em segundo plano."""
    try:
        print("Carregando a planilha principal...")
        planilha = pd.read_excel(arquivo1, engine="openpyxl")
        colunas_existentes = [col for col in planilha.columns if col in COLUNAS_ENTRADA]
        if not colunas_existentes:
            raise ValueError("A planilha principal não possui as colunas esperadas.")
        planilha = planilha[colunas_existentes]
        set_planilha_principal(planilha)  # Salva a planilha no módulo
        print("Planilha principal carregada com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao carregar a planilha principal: {e}")
