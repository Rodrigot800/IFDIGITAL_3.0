import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time
from pacotes.edicaoValorFiltro import  abrir_janela_valores_padroes,valor1,valor2,valor3,valor4
import os
import configparser
config = configparser.ConfigParser()
config.read('config.ini')


# Variáveis globais
planilha_principal = None
planilha_secundaria = None
nomes_vulgares = []  # Lista de todos os nomes vulgares
nomes_selecionados = []  # Lista para manter a ordem dos nomes selecionados






# Colunas de entrada e saída
COLUNAS_ENTRADA = [
    "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
    "X", "Y", "DAP", "Volumes (m³)", "Latitude", "Longitude", "DM", "Observações"
]

COLUNAS_SAIDA = [
    "UT", "Faixa", "Placa", "Nome Vulgar", "Nome Cientifico", "CAP", "ALT", "QF", "X", "Y",
    "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria", "Situacao","UT_AREA_HA_principal"
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
    
    """Carrega a planilha principal e exibe os nomes vulgares."""
    global planilha_principal, nomes_vulgares
    try:
        status_label.config(text="Carregando inventário principal...")
        status_label.pack(pady=10)

        planilha_principal = pd.read_excel(arquivo1, engine="openpyxl")
        print(planilha_principal.columns.tolist())
        colunas_existentes = [col for col in planilha_principal.columns if col in ["Nome Vulgar"]]
        if not colunas_existentes:
            raise ValueError("A planilha principal não possui a coluna 'Nome Vulgar'.")
        nomes_vulgares = sorted(planilha_principal["Nome Vulgar"].dropna().unique()) 
        atualizar_listbox_nomes("")  # Inicializa a Listbox com todos os nomes
        #formatar as colunas 
        planilha_principal.columns = planilha_principal.columns.str.strip()
        planilha_principal.columns = planilha_principal.columns.str.upper() 
        print(planilha_principal.columns.tolist())
        frame_listbox.pack(pady=10)
        frame_secundario.pack(pady=10)
    except Exception as e:
        tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha principal: {e}")
    finally:
        status_label.pack_forget()
        print(planilha_principal.columns)

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

def selecionar_todos():
    """Seleciona todos os nomes vulgares e os adiciona à lista de selecionados."""
    global nomes_selecionados
    nomes_selecionados = nomes_vulgares[:]  # Copia todos os nomes
    atualizar_listbox_selecionados()  # Atualiza a listbox dos selecionados

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


def processar_planilhas(DAPmin,DAPmax,QF,alt):
    

    inicioProcesso = time.time()
    """Processa os dados da planilha principal e mescla com nomes científicos."""
    global planilha_principal, nomes_selecionados

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
            "Observações": "Observacoes",
            "UT_ID" : "UT_ID"
            
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
        df_saida.loc[df_saida["Nome Cientifico"].isna() | (df_saida["Nome Cientifico"]== "") , "Nome Cientifico"] = "NÃO ENCONTRADO"
        df_saida.loc[df_saida["Situacao"].isna() | (df_saida["Situacao"] == ""), "Situacao"] = "SEM RESTRIÇÃO"
        if nomes_selecionados:

            nomes_selecionados = [nome.upper() for nome in nomes_selecionados]

            df_saida["Categoria"] = df_saida["Nome Vulgar"].apply(
            lambda nome: "REM" if nome not in nomes_selecionados else "CORTE"
            )

            def filtrar_REM(row, DAPmin, DAPmax, QF, alt):

                # # Atualizar a coluna "Categoria" com "REM" se a situação for "protegida"
                situacao = str(row["Situacao"]).strip().lower() if pd.notna(row["Situacao"]) else ""

                # Se for protegida, marcar como REM
                if situacao == "protegida":
                    return "REM"

                # Se o DAP for menor que DAPmin ou maior/igual a DAPmax, marcar como REM
                if row["DAP"] < DAPmin or row["DAP"] >= DAPmax:
                    return "REM"

                # Se QF for igual ao valor definido, marcar como REM
                if row["QF"] == QF:
                    return "REM"

                # Se alt > 0 e ALT for maior que alt, marcar como REM
                if alt > 0 and row["ALT"] > alt:
                    return "REM"

                # Se nenhuma condição foi atendida, mantém a categoria original
                return row["Categoria"]

            # Aplicando a função ao DataFrame
            df_saida["Categoria"] = df_saida.apply(lambda row: filtrar_REM(row, DAPmin, DAPmax, QF, alt), axis=1) 
        
        #mesclagem entre ut e ut_area_ha
        
        if "UT" not in planilha_principal.columns or "UT_AREA_HA" not in planilha_principal.columns:
         raise ValueError("As colunas 'UT' ou 'UT_AREA_HA' não foram encontradas no DataFrame planilha_principal.")

        # Garantir que as colunas estejam preenchidas e no tipo correto

        # Verificar valores únicos para depuração
        print("Valores únicos em planilha_principal (UT):", planilha_principal["UT"].unique())
        print("Valores únicos em planilha_principal (UT_ID):", planilha_principal["UT_ID"].unique())

        # Garantir que UT e UT_ID correspondem
        ut_to_area = planilha_principal[["UT", "UT_AREA_HA"]].drop_duplicates().set_index("UT")["UT_AREA_HA"].to_dict()

        # Adicionar UT_AREA_HA diretamente no df_saida
        df_saida["UT_AREA_HA"] = df_saida["UT"].map(ut_to_area)

        print("Valores únicos em planilha_principal (UT):", planilha_principal["UT"].unique())
        print("Valores únicos em planilha_principal (UT_ID):", planilha_principal["UT_ID"].unique())
        print(planilha_principal[["UT_ID", "UT_AREA_HA"]].drop_duplicates())

        print(df_saida[["UT", "UT_ID", "UT_AREA_HA"]].drop_duplicates())


        # Garantir que as colunas relevantes não tenham valores nulos
        planilha_principal["UT"] = planilha_principal["UT"].fillna(0).astype(int)
        planilha_principal["UT_AREA_HA"] = planilha_principal["UT_AREA_HA"].fillna(0).astype(float)
        df_saida["UT"] = df_saida["UT"].fillna(0).astype(int)

        

        # Garantir o mesmo para df_saida
        df_saida["UT_ID"] = df_saida["UT_ID"].fillna(0).astype(int)

        # Realizar a mesclagem com base em UT
        df_saida = pd.merge(
            df_saida,
            planilha_principal[["UT_ID", "UT_AREA_HA"]].drop_duplicates(),
            left_on="UT",
            right_on="UT_ID",
            how="left",
            suffixes=("", "_principal")  # Adiciona "_principal" às colunas duplicadas
        )
        print(df_saida.head())  # Mostrar as primeiras linhas
        print(df_saida[["UT", "UT_ID", "UT_AREA_HA"]].drop_duplicates())  # Verificar valores únicos

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

        print(f"Processamento realizado em {finalProcesso - inicioProcesso:.2f} s e salvo em ")
        tk.messagebox.showinfo("SUCESSO",f" Processamento realizado em {finalProcesso - inicioProcesso:.2f} segundos e arquico salvo em {finalTimeSalvar - inicioTimeSalvar:.2f} s")

    except Exception as e:
        tk.messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")


def iniciar_processamento():
    config = configparser.ConfigParser()
    config.read('config.ini')
    """Inicia o processamento em uma thread separada."""
    thread = threading.Thread(
    target=processar_planilhas,
    args=(
        float(config.get('DEFAULT', 'dapmax')),
        float(config.get('DEFAULT', 'dapmin')),
        int(config.get('DEFAULT', 'qf')),
        float(config.get('DEFAULT', 'alt'))
        )
    )
    thread.daemon = True  # Fecha a thread quando a interface é fechada
    thread.start()

# Função para abrir a janela de valores padrões ao clicar no botão
def abrir_janela_valores_padroes_callback():
    # Abre a janela de valores padrões e bloqueia a janela principal
    janela_padrao = abrir_janela_valores_padroes(app)
    
    # Aguarda até que a janela secundária seja fechada
    app.wait_window(janela_padrao)
    
    

# Interface gráfica
app = tk.Tk()
app.title("IFDIGITAL 3.0")
app.geometry("800x900")


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

botao_modificar_filtro = tk.Button(app, text="Modfificar Filtragem para REM", command=abrir_janela_valores_padroes_callback)
botao_modificar_filtro.pack(pady=5)



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

ttk.Button(frame_listbox, text="Selecionar Todos", command=selecionar_todos).grid(row=3, column=0, pady=10, padx=5)
ttk.Button(frame_listbox, text="Remover Último", command=remover_ultimo_selecionado).grid(row=2, column=0, pady=10)
ttk.Button(frame_listbox, text="Limpar Lista", command=limpar_lista_selecionados).grid(row=2, column=1, pady=10)

# Frame para o botão de processamento
frame_secundario = ttk.Frame(app, padding=(10, 10))
ttk.Button(frame_secundario, text="Processar Planilhas", command=iniciar_processamento, width=40).pack(pady=10)

frame_listbox.pack_forget()  # Inicialmente escondido
frame_secundario.pack_forget()  # Inicialmente escondido

app.mainloop()
