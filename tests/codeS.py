import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time

# Colunas de entrada e saída
COLUNAS_ENTRADA = [
    "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
    "X", "Y", "X Corrigido", "Y Corrigido", "DAP", "Volumes (m³)", "X Negativo", "Y Negativo",
    "Latitude", "Longitude", "DM", "Observações", "N"
]

COLUNAS_SAIDA = [
    "UT", "Faixa", "Placa", "Nome Vulgar", "CAP", "ALT", "QF", "X", "Y",
    "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria", "Nome Cientifico"
]

# Variável para armazenar a planilha carregada
planilha_principal = None

def selecionar_arquivos():
    """Seleciona os arquivos das duas planilhas."""
    global planilha_principal

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
    global planilha_principal
    try:
        print("Carregando a planilha principal...")
        planilha_principal = pd.read_excel(arquivo1, engine="openpyxl")
        colunas_existentes = [col for col in planilha_principal.columns if col in COLUNAS_ENTRADA]
        if not colunas_existentes:
            raise ValueError("A planilha principal não possui as colunas esperadas.")
        planilha_principal = planilha_principal[colunas_existentes]
        print("Planilha principal carregada com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao carregar a planilha principal: {e}")

def processar_planilhas():
    """Processa os dados da planilha principal e mescla com nomes científicos."""
    global planilha_principal

    arquivo2 = entrada2_var.get()

    # Verificar se a planilha principal já foi carregada
    if planilha_principal is None:
        messagebox.showerror("Erro", "Por favor, selecione e aguarde o carregamento da planilha principal.")
        return

    # Verificar se o segundo arquivo foi selecionado
    if not arquivo2:
        messagebox.showerror("Erro", "Por favor, selecione o arquivo de Nomes Vulgares e Científicos.")
        return

    try:
        start_time = time.time()
        progress_var.set(10)  # Atualiza o progresso inicial
        progress_bar.pack(pady=10, fill="x")  # Mostra a barra de progresso
        progress_bar.update()

        # Criar o DataFrame de saída
        print("Criando o DataFrame de saída...")
        df_saida = pd.DataFrame()
        for entrada, saida in {
            "UT": "UT",
            "Faixa": "Faixa",
            "Placa": "Placa",
            "Nome Vulgar": "Nome Vulgar",
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
        }.items():
            if entrada in planilha_principal.columns:
                df_saida[saida] = planilha_principal[entrada]
            else:
                df_saida[saida] = None  # Colunas ausentes preenchidas com None
        if "Categoria" not in df_saida.columns:
            df_saida["Categoria"] = None
        progress_var.set(40)
        progress_bar.update()

        # Carregar a segunda planilha
        print("Carregando a segunda planilha...")
        df2 = pd.read_excel(arquivo2, engine="openpyxl")
        df2.rename(columns={"NOME_VULGAR": "Nome Vulgar", "NOME_CIENTIFICO": "Nome Cientifico"}, inplace=True)

        # Normalizar as colunas "Nome Vulgar" e mesclar os dados
        df_saida["Nome Vulgar"] = df_saida["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Vulgar"] = df2["Nome Vulgar"].str.strip().str.upper()
        df_saida = pd.merge(df_saida, df2[["Nome Vulgar", "Nome Cientifico"]], on="Nome Vulgar", how="left")
        progress_var.set(80)
        progress_bar.update()

        # Reordenar as colunas para garantir que "Nome Cientifico" esteja ao lado de "Nome Vulgar"
        colunas = list(df_saida.columns)
        if "Nome Cientifico" in colunas:
            colunas.remove("Nome Cientifico")
            idx = colunas.index("Nome Vulgar") + 1
            colunas.insert(idx, "Nome Cientifico")
        df_saida = df_saida[colunas]
        
        # Salvar o arquivo de saída
        print("Salvando o arquivo de saída...")
        diretorio = os.path.dirname(entrada1_var.get())
        arquivo_saida = os.path.join(diretorio, "planilha_processada_completa.xlsx")
        df_saida.to_excel(arquivo_saida, index=False, engine="xlsxwriter")
        progress_var.set(100)
        progress_bar.update()

        elapsed_time = time.time() - start_time
        print(f"Processamento concluído em {elapsed_time:.2f} segundos.")
        messagebox.showinfo("Sucesso", f"Planilha processada salva em:\n{arquivo_saida}\nTempo total: {elapsed_time:.2f} segundos")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao processar as planilhas: {e}")
    finally:
        progress_var.set(0)  # Resetar a barra de progresso
        progress_bar.pack_forget()  # Esconde a barra de progresso
        progress_bar.update()

def iniciar_processamento():
    """Inicia o processamento em uma thread separada."""
    thread = threading.Thread(target=processar_planilhas)
    thread.daemon = True  # Fecha a thread quando a interface é fechada
    thread.start()

# Interface gráfica
app = tk.Tk()
app.title("Processador de Inventário e Mesclagem")
app.geometry("600x400")

entrada1_var = tk.StringVar()
entrada2_var = tk.StringVar()
progress_var = tk.IntVar()

tk.Label(app, text="Arquivo 1: Planilha Principal").pack(pady=5)
tk.Entry(app, textvariable=entrada1_var, width=60).pack(pady=5)
tk.Label(app, text="Arquivo 2: Nomes Vulgares e Científicos").pack(pady=5)
tk.Entry(app, textvariable=entrada2_var, width=60).pack(pady=5)

progress_bar = ttk.Progressbar(app, variable=progress_var, maximum=100)

tk.Button(app, text="Selecionar Planilhas", command=selecionar_arquivos).pack(pady=10)
tk.Button(app, text="Processar e Mesclar Planilhas", command=iniciar_processamento, bg="green", fg="white").pack(pady=20)

app.mainloop()
