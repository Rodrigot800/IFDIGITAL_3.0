import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Colunas de entrada e saída
COLUNAS_ENTRADA = [
    "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
    "X", "Y", "X Corrigido", "Y Corrigido", "DAP", "Volumes (m³)", "X Negativo", "Y Negativo",
    "Latitude", "Longitude", "DM", "Observações", "N"
]

COLUNAS_SAIDA = [
    "UT", "Faixa", "Placa", "NomeVulgar", "CAP", "ALT", "QF", "X", "Y",
    "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria"
]

def selecionar_arquivos():
    """Seleciona os arquivos das duas planilhas."""
    arquivo1 = filedialog.askopenfilename(
        title="Selecione a planilha de entrada (dados principais)",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if arquivo1:
        entrada1_var.set(arquivo1)

    arquivo2 = filedialog.askopenfilename(
        title="Selecione a segunda planilha (Nomes Vulgares e Científicos)",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if arquivo2:
        entrada2_var.set(arquivo2)

def processar_planilhas():
    """Processa os dados da planilha principal, mescla com nomes científicos e ajusta a posição da coluna NomeCientifico_y."""
    arquivo1 = entrada1_var.get()
    arquivo2 = entrada2_var.get()

    if not arquivo1 or not arquivo2:
        messagebox.showerror("Erro", "Por favor, selecione ambos os arquivos de entrada.")
        return

    try:
        # Carregar a planilha principal
        df1 = pd.read_excel(arquivo1, engine="openpyxl")

        # Verificar as colunas presentes e remover as que não estão na lista de entrada
        colunas_existentes = [col for col in df1.columns if col in COLUNAS_ENTRADA]
        df1 = df1[colunas_existentes]

        # Criar o DataFrame de saída com as colunas especificadas
        df_saida = pd.DataFrame(columns=COLUNAS_SAIDA)

        # Mapear os valores das colunas de entrada para as de saída
        mapeamento_colunas = {
            "UT": "UT",
            "Faixa": "Faixa",
            "Placa": "Placa",
            "Nome Vulgar": "NomeVulgar",
            "CAP": "CAP",
            "ALT": "ALT",
            "QF": "QF",
            "X": "X",
            "Y": "Y",
            "DAP": "DAP",
            "Volumes (m³)": "Volume_m3",
            "Latitude": "Latitude",
            "Longitude": "Longitude",
            "DM": "DM",
            "Observações": "Observacoes"
        }

        for entrada, saida in mapeamento_colunas.items():
            if entrada in df1.columns:
                df_saida[saida] = df1[entrada]

        # Preencher colunas ausentes com valores nulos
        for coluna in COLUNAS_SAIDA:
            if coluna not in df_saida.columns:
                df_saida[coluna] = None

        # Carregar a segunda planilha
        df2 = pd.read_excel(arquivo2, engine="openpyxl")

        # Verificar se as colunas necessárias estão presentes
        if 'NomeVulgar' not in df2.columns or 'NomeCientifico' not in df2.columns:
            messagebox.showerror(
                "Erro",
                "A segunda planilha deve conter as colunas 'NomeVulgar' e 'NomeCientifico'."
            )
            return

        # Normalizar colunas (remover espaços, acentos e case sensitive)
        df_saida['NomeVulgar'] = df_saida['NomeVulgar'].str.strip().str.upper().str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
        df2['NomeVulgar'] = df2['NomeVulgar'].str.strip().str.upper().str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

        # Relacionar as planilhas com base no NomeVulgar
        df_saida = pd.merge(df_saida, df2[['NomeVulgar', 'NomeCientifico']], on='NomeVulgar', how='left')

        # Reorganizar a coluna NomeCientifico para ficar ao lado de NomeVulgar
        colunas_reordenadas = list(df_saida.columns)
        colunas_reordenadas.remove("NomeCientifico")
        index_nome_vulgar = colunas_reordenadas.index("NomeVulgar")
        colunas_reordenadas.insert(index_nome_vulgar + 1, "NomeCientifico")
        df_saida = df_saida[colunas_reordenadas]

        # Criar um novo arquivo no mesmo diretório
        diretorio = os.path.dirname(arquivo1)
        arquivo_saida = os.path.join(diretorio, "planilha_processada_completa.xlsx")

        # Salvando a planilha final
        df_saida.to_excel(arquivo_saida, index=False, engine="openpyxl")

        messagebox.showinfo("Sucesso", f"Planilha processada salva em:\n{arquivo_saida}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao processar as planilhas: {e}")

# Interface gráfica
app = tk.Tk()
app.title("Processador de Inventário e Mesclagem")
app.geometry("600x400")

# Variáveis para os arquivos
entrada1_var = tk.StringVar()
entrada2_var = tk.StringVar()

# Componentes da interface
tk.Label(app, text="Arquivo 1: Planilha Principal").pack(pady=5)
tk.Entry(app, textvariable=entrada1_var, width=60).pack(pady=5)
tk.Label(app, text="Arquivo 2: Nomes Vulgares e Científicos").pack(pady=5)
tk.Entry(app, textvariable=entrada2_var, width=60).pack(pady=5)
tk.Button(app, text="Selecionar Planilhas", command=selecionar_arquivos).pack(pady=10)
tk.Button(app, text="Processar e Mesclar Planilhas", command=processar_planilhas, bg="green", fg="white").pack(pady=20)

# Rodar a interface
app.mainloop()
