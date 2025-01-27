import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time
from pacotes.edicaoValorFiltro import abrir_janela_valores_padroes

# Variáveis globais
planilha_principal = None
planilha_secundaria = None
nomes_vulgares = []  # Lista de todos os nomes vulgares
nomes_selecionados = []  # Lista para manter a ordem dos nomes selecionados
start_total = None

# Colunas de entrada e saída
COLUNAS_ENTRADA = [
    "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
    "X", "Y", "DAP", "Volumes (m³)", "Latitude", "Longitude", "DM", "Observações"
]

COLUNAS_SAIDA = [
    "UT", "Faixa", "Placa", "Nome Vulgar", "Nome Cientifico", "CAP", "ALT", "QF", "X", "Y",
    "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria", "Situacao"
]

# Funções da Interface e Processamento

def selecionar_arquivos(tipo):
    """Seleciona os arquivos das planilhas."""
    global planilha_principal, planilha_secundaria

    arquivo = filedialog.askopenfilename(
        title=f"Selecione a planilha {'principal' if tipo == 'principal' else 'secundária'}",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if arquivo:
        if tipo == "principal":
            entrada1_var.set(arquivo)
            threading.Thread(target=carregar_planilha_principal, args=(arquivo,)).start()
        elif tipo == "secundaria":
            entrada2_var.set(arquivo)
            threading.Thread(target=carregar_planilha_secundaria, args=(arquivo,)).start()


def carregar_planilha_principal(arquivo1):
    global start_total
    start_total = time.time() 
    start_part1 = time.time()
    """Carrega a planilha principal e exibe os nomes vulgares."""
    global planilha_principal, nomes_vulgares
    try:
        status_label.config(text="Carregando inventário principal...")
        status_label.pack(pady=10)

        planilha_principal = pd.read_excel(arquivo1, engine="openpyxl")
        colunas_existentes = [col for col in planilha_principal.columns if col in ["Nome Vulgar"]]
        if not colunas_existentes:
            raise ValueError("A planilha principal não possui a coluna 'Nome Vulgar'.")
        nomes_vulgares = sorted(planilha_principal["Nome Vulgar"].dropna().unique()) 
        atualizar_listbox_nomes("")  # Inicializa a Listbox com todos os nomes

        frame_listbox.pack(pady=10)
        frame_secundario.pack(pady=10)
    except Exception as e:
        tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha principal: {e}")
    finally:
        status_label.pack_forget()
        end_part1 = time.time()
        print(f"Tempo para carregar a planilha principal: {end_part1 - start_part1:.2f} segundos")

def carregar_planilha_secundaria(arquivo2):
    """Carrega a planilha secundária em segundo plano."""
    global planilha_secundaria
    try:
        planilha_secundaria = pd.read_excel(arquivo2, engine="openpyxl")
        print("Planilha secundária carregada com sucesso.")
    except Exception as e:
        tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha secundária: {e}")


def atualizar_listbox_nomes(filtro):
    """Atualiza a Listbox com nomes vulgares que atendem ao filtro."""
    listbox_nomes_vulgares.delete(0, tk.END)
    for nome in nomes_vulgares:
        if filtro.lower() in nome.lower():
            listbox_nomes_vulgares.insert(tk.END, nome)


def atualizar_listbox_selecionados():
    """Atualiza a Listbox com os nomes selecionados."""
    listbox_selecionados.delete(0, tk.END)
    for nome in nomes_selecionados:
        listbox_selecionados.insert(tk.END, nome)


def pesquisar_nomes(event):
    """Callback para filtrar nomes com base na pesquisa."""
    filtro = pesquisa_var.get()
    atualizar_listbox_nomes(filtro)


def adicionar_selecao(event):
    """Adiciona um nome à lista de selecionados ao clicar."""
    global nomes_selecionados

    selecao = listbox_nomes_vulgares.curselection()
    if selecao:
        nome = listbox_nomes_vulgares.get(selecao[0])  # Obtém o nome selecionado
        if nome not in nomes_selecionados:
            nomes_selecionados.append(nome)
            atualizar_listbox_selecionados()


def remover_ultimo_selecionado():
    """Remove o último nome adicionado à lista de selecionados."""
    if nomes_selecionados:
        nomes_selecionados.pop()
        atualizar_listbox_selecionados()


def limpar_lista_selecionados():
    """Limpa todos os nomes da lista de selecionados."""
    nomes_selecionados.clear()
    atualizar_listbox_selecionados()


def processar_planilhas():
    inicioProcesso = time.time()
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
        # Código do processamento
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
                df_saida[saida] = None
            if "Categoria" not in df_saida.columns:
                df_saida["Categoria"] = None
        
        df2 = pd.read_excel(arquivo2,engine="openpyxl")
        df2.rename(columns={
            "NOME_VULGAR": "Nome Vulgar",
            "NOME_CIENTIFICO": "Nome Cientifico",
            "SITUACAO":"Situacao"
            }, inplace=True)
        
        df_saida["Nome Vulgar"] = df_saida["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Vulgar"] = df2["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Cientifico"] = df2["Nome Cientifico"].str.strip().str.upper()

        df_saida = pd.merge(df_saida,df2[["Nome Vulgar","Nome Cientifico","Situacao"]],
                            on="Nome Vulgar", how="left")
        df_saida.loc[df_saida["Situacao"].str.lower() == "protegida", "Categoria"] = "REM"
        df_saida.loc[df_saida["DAP"] < 0.5, "Categoria"] = "REM"
        df_saida.loc[df_saida["DAP"] >= 2, "Categoria"] = "REM"
        df_saida.loc[df_saida["QF"] == 3, "Categoria"] = "REM"
        df_saida.loc[df_saida["Nome Cientifico"].isna() | (df_saida["Nome Cientifico"]== "") , "Nome Cientifico"] = "Não encontrado"

        #organizando as colunas
        df_saida = df_saida[COLUNAS_SAIDA]


        finalProcesso = time.time()
        print(f"Processamento realizado em {finalProcesso - inicioProcesso:.2f} s")

        inicioTimeSalvar = time.time()
        #salvar o arquivo no diretório
        diretorio = os.path.dirname(entrada1_var.get())
        arquivo_saida = os.path.join(diretorio, "Planilha Processada - IFDIGITAL 3.0.xlsx")
        df_saida.to_excel(arquivo_saida,index=False,engine="xlsxwriter")
        finalTimeSalvar = time.time()
        print(f" arquico salvo em {finalTimeSalvar - inicioTimeSalvar:.2f} s")

        end_total = time.time()
        tempoTotal = end_total - start_total
        print(f"Processamento realizado em {tempoTotal:.2f} s")
        tk.messagebox.showinfo("SUCESSO",f" Processamento realizado em {tempoTotal:.2f} segundos")

    except Exception as e:
        tk.messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")


def iniciar_processamento():
    """Inicia o processamento em uma thread separada."""
    thread = threading.Thread(target=processar_planilhas)
    thread.daemon = True  # Fecha a thread quando a interface é fechada
    thread.start()


# Interface gráfica
app = tk.Tk()
app.title("IFDIGITAL 3.0")
app.geometry("900x700")

entrada1_var = tk.StringVar()
entrada2_var = tk.StringVar()
pesquisa_var = tk.StringVar()

# Frame para entrada de arquivos
frame_inputs = ttk.LabelFrame(app, text="Entrada de Arquivos", padding=(10, 10))
frame_inputs.pack(fill="x", pady=10, padx=10)

ttk.Label(frame_inputs, text="Arquivo 1: Planilha Principal").grid(row=0, column=0, sticky="w")
ttk.Entry(frame_inputs, textvariable=entrada1_var, width=60).grid(row=0, column=1, pady=5, padx=5)
ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("principal")).grid(row=0, column=2, padx=5)

ttk.Label(frame_inputs, text="Arquivo 2: Planilha Secundária").grid(row=1, column=0, sticky="w")
ttk.Entry(frame_inputs, textvariable=entrada2_var, width=60).grid(row=1, column=1, pady=5, padx=5)
ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("secundaria")).grid(row=1, column=2, padx=5)
#ttk.Button(frame_inputs, text="Novo Botão", command=lambda:abrir_janela_valores_padroes(root) ).grid(row=2, column=1, pady=10, padx=5)  # Botão abaixo dos inputs
botao_modificar_filtro = tk.Button(app, text="Novo Botão", command=lambda:abrir_janela_valores_padroes )
botao_modificar_filtro.pack(pady=5)

frame_dados_de_Filtragem = ttk.LabelFrame(app, text="Dados de Filtragem", padding=(10,10))
frame_dados_de_Filtragem.pack(fill="x", pady=10, padx=10)

status_label = ttk.Label(app, text="")

# Frame para Listboxes
frame_listbox = ttk.LabelFrame(app, text="Seleção de Nomes Vulgares", padding=(10, 10))

ttk.Label(frame_listbox, text="Pesquisar:").grid(row=0, column=0, sticky="w")
pesquisa_entry = ttk.Entry(frame_listbox, textvariable=pesquisa_var, width=40)
pesquisa_entry.grid(row=0, column=1, padx=10, pady=5)
pesquisa_entry.bind("<KeyRelease>", pesquisar_nomes)

listbox_nomes_vulgares = tk.Listbox(frame_listbox, selectmode=tk.SINGLE, width=40, height=20)
listbox_nomes_vulgares.bind("<<ListboxSelect>>", adicionar_selecao)
listbox_nomes_vulgares.grid(row=1, column=0, padx=10, pady=10)

listbox_selecionados = tk.Listbox(frame_listbox, width=40, height=20)
listbox_selecionados.grid(row=1, column=1, padx=10, pady=10)

ttk.Button(frame_listbox, text="Remover Último", command=remover_ultimo_selecionado).grid(row=2, column=0, pady=10)
ttk.Button(frame_listbox, text="Limpar Lista", command=limpar_lista_selecionados).grid(row=2, column=1, pady=10)

# Frame para o botão de processamento
frame_secundario = ttk.Frame(app, padding=(10, 10))
ttk.Button(frame_secundario, text="Processar Planilhas", command=iniciar_processamento, width=40).pack(pady=10)

frame_listbox.pack_forget()  # Inicialmente escondido
frame_secundario.pack_forget()  # Inicialmente escondido

app.mainloop()
