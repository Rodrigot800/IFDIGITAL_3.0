import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time
from pacotes.edicaoValorFiltro import  abrir_janela_valores_padroes,valor1,valor2,valor3,valor4
from pacotes.ordemSubstituta import OrdenadorFrame 
import os
import sys
import configparser
import numpy as np
from PIL import Image, ImageTk
from tkinter import PhotoImage


config = configparser.ConfigParser()
config.read('config.ini')
CONFIG_FILE = 'config.ini'

# Variáveis globais
planilha_principal = None
planilha_secundaria = None
nomes_vulgares = []  # Lista de todos os nomes vulgares
especies_selecionados = []  # Lista para manter a ordem dos nomes selecionados
nomes_selecionados  = []
start_total = None
ordering_mode = "QF > Vol"

df_tabelaDeAjusteVol = None
global df_valores_atualizados
global  df_saida 
dados_editados_por_ut = {}

# Flag que indica se a tabela está visível ou não
tabela_visivel = True


# Colunas de entrada e saída
COLUNAS_ENTRADA = [
    "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
    "X", "Y", "DAP", "Volumes (m³)", "Latitude", "Longitude", "DM", "Observações"
]

COLUNAS_SAIDA = [
    "UT", "Faixa", "Placa", "Nome Vulgar", "Nome Cientifico", "CAP", "H", "QF", "X", "Y",
    "DAP", "Vol", "Lat", "Long", "DM", "OBS", "Categoria", "Situacao","UT_AREA_HA"
]



# Funções da Interface e Processamento

def update_ordering_mode(event=None):
    """
    Atualiza a variável global 'ordering_mode' com o valor selecionado no combobox 
    e ordena o DataFrame automaticamente.
    """
    global ordering_mode
    ordering_mode = combobox.get()
    print("Modo de ordenação atualizado para:", ordering_mode)

def salvar_caminho(tipo, caminho):
    """Salva o caminho da planilha no arquivo de configuração."""
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)
    if not config.has_section("Planilha"):
        config.add_section("Planilha")
    config.set("Planilha", tipo, caminho)
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

def selecionar_arquivos(tipo):
    """
    Seleciona os arquivos das planilhas e salva o caminho.
    Caso o usuário selecione um arquivo, atualiza o widget de entrada e inicia o carregamento.
    """
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
        salvar_caminho(tipo, arquivo)
        
def carregar_planilha_salva(tipo):
    """
    Verifica se há um caminho salvo para a planilha do tipo especificado e,
    se existir, atualiza o widget de entrada e inicia o carregamento em uma thread.
    Essa função pode ser chamada na inicialização do projeto para carregar automaticamente os caminhos salvos.
    """
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)
    caminho_salvo = config.get("Planilha", tipo, fallback="")
    if caminho_salvo:
        if tipo == "principal":
            entrada1_var.set(caminho_salvo)
            threading.Thread(target=carregar_planilha_principal, args=(caminho_salvo,)).start()
        elif tipo == "secundaria":
            entrada2_var.set(caminho_salvo)
            threading.Thread(target=carregar_planilha_secundaria, args=(caminho_salvo,)).start()


def carregar_planilha_principal(arquivo1):
    
    """Carrega a planilha principal e exibe os nomes vulgares."""
    global planilha_principal, nomes_vulgares
    try:
        status_label.config(text="Carregando inventário principal...")
        status_label.place(x= 700,y=150)

        planilha_principal = pd.read_excel(arquivo1, engine="openpyxl")
        colunas_existentes = [col for col in planilha_principal.columns if col in ["Nome Vulgar"]]
        if not colunas_existentes:
            raise ValueError("A planilha principal não possui a coluna 'Nome Vulgar'.")
        nomes_vulgares = sorted(planilha_principal["Nome Vulgar"].dropna().unique()) 
        atualizar_listbox_nomes("")  # Inicializa a Listbox com todos os nomes

        frame_listbox_e_tabela.place(relx=0.05, rely=0.2)
        frame_secundario.place(relx=0.425, rely=0.90)

    except Exception as e:
        tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha principal: {e}")
    finally:
        status_label.place_forget()

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


def pesquisar_nomes(event):
    """Callback para filtrar nomes com base na pesquisa."""
    filtro = pesquisa_var.get()
    atualizar_listbox_nomes(filtro)

# Função para adicionar ao clicar na Listbox
def adicionar_selecao(event):
    indices = listbox_nomes_vulgares.curselection()
    if not indices:
        return
    
    index = indices[0]
    nome = listbox_nomes_vulgares.get(index)

    # Verifica duplicatas antes de adicionar
    for child in table_selecionados.get_children():
        valores = table_selecionados.item(child, "values")
        if valores and valores[0] == nome:
            return  # Já está na tabela, não adiciona novamente

    # Recarregar os valores do config.ini ANTES de adicionar
    config.read("config.ini")  # Garante que estamos lendo a versão mais recente do arquivo

    dap_max = config.getfloat("DEFAULT", "dapmax", fallback=0.5)
    dap_min = config.getfloat("DEFAULT", "dapmin", fallback=5)
    qf = config.getint("DEFAULT", "qf", fallback=3)
    alt = config.getfloat("DEFAULT", "alt", fallback=0)
    cap = config.getfloat("DEFAULT", "cap", fallback=2.5)

    # Se alt for 0, substituir por string vazia
    alt = "" if alt == 0 else alt

    # Alternando as cores das linhas
    tag = 'verde_claro' if  len(table_selecionados.get_children()) % 2 == 0 else 'branca'
    
    # Adiciona à tabela com os valores atualizados do config.ini
    valores_atualizados = (nome, dap_max, dap_min, qf, alt, cap)
    table_selecionados.insert("", "end", values=valores_atualizados, tags=(tag,))

    # Exemplo de como você pode adicionar as tags de cores à tabela
    table_selecionados.tag_configure('branca', background='white')
    table_selecionados.tag_configure('verde_claro', background='#d3f8e2')

def editar_linha(event):
    selected_item = table_selecionados.focus()
    if not selected_item:
        return

    valores = table_selecionados.item(selected_item, "values")
    if not valores:
        return

    def salvar_edicao():
        novos_valores = [valores[0]] + [entry.get() for entry in entradas]
        table_selecionados.item(selected_item, values=novos_valores)
        popup.destroy()

    popup = tk.Toplevel(app)
    popup.title("Editar Linha")

    ttk.Label(popup, text="Nome da Espécie:").grid(row=0, column=0, padx=5, pady=5)
    ttk.Label(popup, text=valores[0], font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=5)

    entradas = []
    for i, coluna in enumerate(colunas_selecionados[1:], start=1):
        ttk.Label(popup, text=coluna).grid(row=i, column=0, padx=5, pady=5)
        entry = ttk.Entry(popup)
        entry.grid(row=i, column=1, padx=5, pady=5)
        entry.insert(0, valores[i])
        entradas.append(entry)

    ttk.Button(popup, text="Salvar", command=salvar_edicao).grid(row=len(colunas_selecionados), column=0, columnspan=2, pady=10)

# Função para selecionar todos os itens da Listbox
def selecionar_todos():
    # Carrega os valores do config.ini
    config.read("config.ini")  # Garante que estamos lendo a versão mais recente do arquivo

    dap_max = config.getfloat("DEFAULT", "dapmax", fallback=0.5)
    dap_min = config.getfloat("DEFAULT", "dapmin", fallback=5)
    qf = config.getint("DEFAULT", "qf", fallback=3)
    alt = config.getfloat("DEFAULT", "alt", fallback=0)

    # Percorre todos os nomes vulgares e adiciona à tabela se não existirem
    for nome in nomes_vulgares:
        # Verifica se o nome já existe na tabela para evitar duplicação
        is_duplicate = any(table_selecionados.item(child, "values")[0] == nome for child in table_selecionados.get_children())
        
        if not is_duplicate:  # Só adiciona se o nome não estiver na tabela
            # Adiciona os dados na tabela com valores do config.ini
            valores_atualizados = (nome, dap_max, dap_min, qf, alt)
            table_selecionados.insert("", "end", values=valores_atualizados)
            print(f"{nome} foi adicionado à tabela.")  # Mensagem de depuração para verificar se está sendo adicionado
        else:
            print(f"{nome} já está na tabela.")  # Mensagem de depuração para duplicação

    print("Todos os itens de 'nomes_vulgares' foram processados.")

# Função para remover o último item da tabela
def remover_ultimo_selecionado():
    filhos = table_selecionados.get_children()
    if filhos:
        table_selecionados.delete(filhos[-1])

# Função para limpar todos os itens da tabela
def limpar_lista_selecionados():
    table_selecionados.delete(*table_selecionados.get_children())

def tabelaDeResumo():
     # Exemplo de dados (você pode substituir com os seus)
    dados = [
        {
            "especie": "Bertholletia excelsa",
            "n_arvores": 50,
            "vol_arvore": 0.8,
            "vol_ajustado": 40,
            "vol_arvore_ajustado": 0.7,
            "vol_ha": 5.5,
            "cap_min": 45
        },
        {
            "especie": "Swietenia macrophylla",
            "n_arvores": 30,
            "vol_arvore": 1.2,
            "vol_ajustado": 36,
            "vol_arvore_ajustado": 1.0,
            "vol_ha": 4.2,
            "cap_min": 50
        }
    ]

    # Limpa a tabela antes de inserir novos dados
    for item in table_resumo_especie.get_children():
        table_resumo_especie.delete(item)

    # Insere os dados na tabela
    for item in dados:
        table_resumo_especie.insert("", "end", values=(
            item["especie"],
            item["n_arvores"],
            item["vol_arvore"],
            item["vol_ajustado"],
            item["vol_arvore_ajustado"],
            item["vol_ha"],
            item["cap_min"]
        ))

def ajustarVolumeHect():
    global df_tabelaDeAjusteVol, df_valores_atualizados,df_saida
    

    # Cria uma cópia do DataFrame original para não modificá-lo
    df_modificado = df_saida.copy()

    # Cria as colunas mCAP e mH com valores padrão: mCAP = 1, mH = 0
    df_modificado["mCAP"] = 1
    df_modificado["mH"] = 0

    if 'df_valores_atualizados' in globals() and not df_valores_atualizados.empty:
        for col in ["mCAP", "mH"]:
            # Buscar os valores originais com base no novo nome
            nome_origem = col[1:]  # remove o 'm' do início
            df_valores_atualizados[col] = df_valores_atualizados.get(nome_origem, df_valores_atualizados.get(col))
        
        df_modificado = df_modificado.merge(
            df_valores_atualizados[["ut", "mCAP", "mH"]],
            left_on="UT", right_on="ut", how="left", suffixes=("", "_novo")
        )
        for col in ["mCAP", "mH"]:
            if f"{col}_novo" in df_modificado.columns:
                novo_val = df_modificado[f"{col}_novo"]
                df_modificado[col] = np.where(novo_val.notna(), novo_val, df_modificado[col])
        
        cols_to_drop = [f"{col}_novo" for col in ["mCAP", "mH"]]
        df_modificado.drop(columns=cols_to_drop, inplace=True)
        if "ut" in df_modificado.columns and "UT" in df_modificado.columns:
            df_modificado.drop(columns=["ut"], inplace=True)


    df_modificado["ut"] = df_modificado["UT"]
    df_modificado["hac"] = df_modificado["UT_AREA_HA"]

    # Calculando a coluna 'H_a' dependendo da condição de 'Categoria'
    df_modificado["H_a"] = df_modificado.apply(
        lambda row: row["H"] + row["mH"] if row["Categoria"] == "CORTE" else row["H"],
        axis=1
    )
    
    # Calculando a coluna 'CAP_a' somente para linhas com 'Categoria' igual a 'CORTE'
    df_modificado["CAP_a"] = df_modificado.apply(
        lambda row: round(row["CAP"] * row["mCAP"]) if row["Categoria"] == "CORTE" else row["CAP"],
        axis=1
    )

    df_modificado["DAP_a"] = df_modificado.apply(
        lambda row: ((row["CAP_a"] / np.pi) / 100) if row["Categoria"] == "CORTE" else row["DAP"],
        axis=1
    )

    # Aplicando as condições em 'H_a' somente para linhas com 'Categoria' igual a 'CORTE'
    df_modificado["H_a"] = df_modificado.apply(
        lambda row: 10 if row["H"] > 10 and row["H_a"] <= 10 else 
                    row["H_a"] if row["H"] > 10 else 
                    row["H"] if row["H_a"] < row["H"] else row["H_a"]
                    if row["Categoria"] == "CORTE" else row["H_a"],  # Caso não seja CORTE, mantém o valor original
        axis=1
    )


    # Aplicando as condições em 'CAP_a' somente para linhas com 'Categoria' igual a 'CORTE'
    df_modificado["CAP_a"] = df_modificado.apply(
        lambda row: 158 if row["CAP"] > 158 and row["CAP_a"] <= 158 else  
                    row["CAP_a"] if row["CAP"] > 158 else  
                    row["CAP"] if row["CAP_a"] < row["CAP"] else 
                    min(row["CAP_a"], 628) if row["CAP_a"] >= 680 else row["CAP_a"]
                    if row["Categoria"] == "CORTE" else row["CAP_a"],  # Caso não seja CORTE, mantém o valor original
        axis=1
    )

    # Calculando a coluna 'DAP_a' somente para linhas com 'Categoria' igual a 'CORTE'
    df_modificado["DAP_a"] = df_modificado.apply(
        lambda row: ((row["CAP_a"] / np.pi) / 100) 
            if row["Categoria"] == "CORTE" else row["DAP"],  # Caso não seja CORTE, mantém o valor original
        axis=1
    )

    df_modificado["Vol_a"] = ((df_modificado["DAP_a"] ** 2) * np.pi / 4) * df_modificado["H_a"] * 0.7

    df_modificado["Categoria"] = df_modificado["Categoria"].astype(str).str.strip().str.upper()

    df_filtrado = df_modificado[df_modificado["Categoria"].isin(["CORTE"])]

    if df_filtrado.empty:
        print("Nenhuma árvore foi categorizada como 'CORTE'.")
        df_tabelaDeAjusteVol = df_modificado[["ut", "hac", "mCAP", "mH", "DAP_a"]].drop_duplicates()
        df_tabelaDeAjusteVol["num_arvores"] = 0
        df_tabelaDeAjusteVol["volume_total"] = 0
        df_tabelaDeAjusteVol["diminuir"] = 0
        df_tabelaDeAjusteVol["aumentar"] = 0
        df_tabelaDeAjusteVol["media_dif_dap"] = 0
    else:
        # Agrupa para somar DAP e DAP_a por UT
        df_soma_dap = df_filtrado.groupby("ut")[["DAP", "DAP_a"]].sum().reset_index()

        # Calcula a diferença percentual entre DAP_a e DAP
        df_soma_dap["dif_pct"] = ((df_soma_dap["DAP_a"] - df_soma_dap["DAP"]) / df_soma_dap["DAP"]) * 100

        # Agrupa para contar o número de árvores e somar o volume total por UT
        contagem_arvores = df_filtrado.groupby("ut").size().reset_index(name="num_arvores")
        volume_total_por_ut = df_filtrado.groupby("ut")["Vol_a"].sum().reset_index()
        volume_total_por_ut.rename(columns={"Vol_a": "volume_total"}, inplace=True)
        
        # Cria o DataFrame de ajuste com UT, Hectares, mCAP e mH
        df_tabelaDeAjusteVol = df_modificado[["ut", "hac", "mCAP", "mH", "DAP_a"]].drop_duplicates()
        print("Dados iniciais de UT, Hectares, mCAP, mH e DAP_a:")
        print(df_tabelaDeAjusteVol)

        # Calcula o volume máximo como (hectares * 30)
        df_tabelaDeAjusteVol["volume_max"] = df_tabelaDeAjusteVol["hac"] * 30

        # Faz merge para incorporar a contagem de árvores, o volume total e a diferença percentual
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(contagem_arvores, on="ut", how="left").fillna(0)
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(volume_total_por_ut, on="ut", how="left").fillna(0)
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(df_soma_dap[["ut", "dif_pct"]], on="ut", how="left").fillna(0)

        # Calcula o volume por hectare (V³/ha)
        df_tabelaDeAjusteVol["volume_por_hectare"] = df_tabelaDeAjusteVol.apply(
            lambda row: row["volume_total"] / row["hac"] if row["hac"] > 0 else 0, axis=1
        )
        
        # Calcula as diferenças para "diminuir" e "aumentar"
        df_tabelaDeAjusteVol["diminuir"] = df_tabelaDeAjusteVol.apply(
            lambda row: row["volume_total"] - row["volume_max"] if row["volume_total"] > row["volume_max"] else 0, axis=1
        )
        df_tabelaDeAjusteVol["aumentar"] = df_tabelaDeAjusteVol.apply(
            lambda row: row["volume_max"] - row["volume_total"] if row["volume_total"] < row["volume_max"] else 0, axis=1
        )

        # Remover duplicatas por UT após os merges
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.drop_duplicates(subset=["ut"])

    df_tabelaDeAjusteVol["num_arvores"] = df_tabelaDeAjusteVol["num_arvores"].astype(int)

   # Definir as tags para alternar entre as cores
    table_ut_vol.tag_configure('branca', background='white')  # Cor para as linhas brancas
    table_ut_vol.tag_configure('verde_claro', background='#d3f8e2')  # Cor para as linhas verde claro

    # Limpa a Treeview table_ut_vol para evitar duplicação
    for child in table_ut_vol.get_children():
        table_ut_vol.delete(child)

    # Insere os valores na Treeview table_ut_vol com alternância de cores
    for i, (_, row) in enumerate(df_tabelaDeAjusteVol.iterrows()):
        # Formatando os valores para exibição
        ut_val = f"{row['ut']:.0f}"
        hectares_val = f"{row['hac']:.5f}"
        num_arvores = f"{row['num_arvores']:.0f}"
        volume_total = f"{row['volume_total']:.5f}"
        volume_max = f"{row['volume_max']:.5f}"
        volume_por_hect = f"{row['volume_por_hectare']:.3f}"
        diminuir_val = f"{row['diminuir']:.3f}"
        aumentar_val = f"{row['aumentar']:.3f}"
        dif_pct_val = f"{row['dif_pct']:.2f}"  # Exibe a diferença percentual
        CAP_val = f"{row['mCAP']:.3f}"
        ALT_val = f"{row['mH']:.2f}"

        # Alterna as cores das linhas (índices pares e ímpares)
        tag = 'verde_claro' if i % 2 == 0 else 'branca'
        
        # Insere os dados na Treeview com a tag de cor
        table_ut_vol.insert("", "end", values=(ut_val, hectares_val, num_arvores,
                                                volume_total, volume_max, diminuir_val,
                                                aumentar_val, volume_por_hect,
                                                dif_pct_val, CAP_val, ALT_val), tags=(tag,))

    print("UT, Hectares, Número de Árvores e Volume Total Atualizados:")

    # Retorna o DataFrame modificado
    return df_modificado


def processar_planilhas(save):
    
     # Oculta o botão e exibe/inicia a barra de progresso
    button_process.pack_forget()
    progress_bar.pack(pady=0)
    progress_bar.start(10)
    app.update_idletasks()  # Garante que a interface seja atualizada


    inicioProcesso = time.time()
    """Processa os dados da planilha principal e mescla com nomes científicos."""
    global planilha_principal, especies_selecionados, df_saida

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
            "ALT": "H",
            "QF": "QF",
            "X": "X",
            "Y": "Y",
            "DAP": "DAP",
            "Volumes (m³)": "Vol",
            "Latitude": "Lat",
            "Longitude": "Long",
            "DM": "DM",
            "Observações": "OBS",
            "UT_AREA_HA" : "UT_AREA_HA",
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

        df2 = df2.drop_duplicates(subset=["Nome Vulgar"])
        df_saida = pd.merge(df_saida,df2[["Nome Vulgar","Nome Cientifico","Situacao"]],
                            on="Nome Vulgar", how="left")
        df_saida.loc[df_saida["Nome Cientifico"].isna() | (df_saida["Nome Cientifico"]== "") , "Nome Cientifico"] = "NÃO ENCONTRADO"
        df_saida.loc[df_saida["Situacao"].isna() | (df_saida["Situacao"] == ""), "Situacao"] = "SEM RESTRIÇÃO"


        def extrair_nomes_especies():
            nomes_especies = []
            for child in table_selecionados.get_children():
                valores = table_selecionados.item(child, "values")
                if valores and len(valores) > 0:
                    nomes_especies.append(valores[0].upper())  # Apenas os nomes, convertidos para maiúsculas
            return nomes_especies

        nomes_selecionados = extrair_nomes_especies()
        print(nomes_selecionados) 

       

        # Primeiro, marque como "REMANESCENTE" se a espécie não estiver selecionada.
        df_saida["Categoria"] = df_saida["Nome Vulgar"].apply(
            lambda nome: "REMANESCENTE" if nome.upper() not in nomes_selecionados else "CORTE"
        )

        # Função para extrair valores da tabela
        def extrair_parametros_tabela():
            parametros = {}
            for child in table_selecionados.get_children():
                valores = table_selecionados.item(child, "values")
                if valores and len(valores) > 0:
                    nome = valores[0].upper()
                    # Adiciona os parâmetros da espécie no dicionário com conversão para float
                    try:
                        parametros[nome] = {
                            "dap_min": float(valores[1]) if valores[1] else 0.0,
                            "dap_max": float(valores[2]) if valores[2] else 0.0,
                            "qf": int(valores[3]) if valores[3] else 0,
                            "alt": float(valores[4]) if valores[4] else 0.0
                        }
                    except ValueError as e:
                        print(f"Erro ao processar valores para {nome}: {e}")
                        continue
            return parametros

        # Função para aplicar o filtro
        def filtrar_REM(row, parametros):
            nome = str(row["Nome Vulgar"]).upper()
            ut = str(row["UT"])

            if nome not in parametros:
                return row["Categoria"]

            especie_parametros = parametros[nome]
            DAPmin = especie_parametros["dap_min"]
            DAPmax = especie_parametros["dap_max"]
            QF = especie_parametros["qf"]
            alt = especie_parametros["alt"]

            # Verifica se há edições salvas para a UT e espécie
            if ut in dados_editados_por_ut and nome in dados_editados_por_ut[ut]:
                edits = dados_editados_por_ut[ut][nome]

                # Se F_REM for "SIM", marca como REMANESCENTE diretamente
                if edits.get("F_REM", "").strip().upper() == "SIM":
                    return "REMANESCENTE"

                # Substitui DAPmin e QF se estiverem definidos
                if "DAP <" in edits:
                    try:
                        DAPmin = float(edits["DAP <"])
                    except ValueError:
                        pass

                if "QF >=" in edits:
                    try:
                        QF = int(edits["QF >="])
                    except ValueError:
                        pass

            # Protegidas continuam sendo REMANESCENTE
            situacao = str(row.get("Situacao", "")).strip().lower()
            if situacao == "protegida":
                return "REMANESCENTE"

            # Regras padrão
            if isinstance(row["DAP"], (float, int)) and (row["DAP"] < DAPmin or row["DAP"] >= DAPmax):
                return "REMANESCENTE"

            if isinstance(row["QF"], int) and row["QF"] >= QF:
                return "REMANESCENTE"

            if isinstance(row["H"], (float, int)) and alt > 0 and row["H"] > alt:
                return "REMANESCENTE"

            return row["Categoria"]

        # Recuperar os parâmetros da tabela
        parametros_tabela = extrair_parametros_tabela()

        # Aplicando a função ao DataFrame
        df_saida["Categoria"] = df_saida.apply(lambda row: filtrar_REM(row, parametros_tabela), axis=1)


        df_saida = pd.merge(
            df_saida,
            planilha_principal[["UT_ID", "UT_AREA_HA"]].drop_duplicates(),
            left_on="UT",
            right_on="UT_ID",
            how="left",
            suffixes=("", "_principal")
        )
        # Imprime o DataFrame final para confirmação



        # Atualizar as colunas para evitar duplicações
        df_saida["UT_ID"] = df_saida["UT_ID_principal"]
        df_saida["UT_AREA_HA"] = df_saida["UT_AREA_HA_principal"]

        # Remover as colunas duplicadas
        df_saida.drop(columns=["UT_ID_principal", "UT_AREA_HA_principal"], inplace=True)

        # Verificar o resultado final
        #print(df_saida[["UT", "UT_ID", "UT_AREA_HA"]].drop_duplicates())
        
        #indice de raridade e classificação de substituta 
        ####----------------------

        # **Filtrar apenas as espécies selecionadas que também estão classificadas como "CORTE"**
        df_filtrado = df_saida[
            (df_saida["Nome Vulgar"].isin(nomes_selecionados)) &
            (df_saida["Categoria"] == "CORTE")  # Apenas árvores marcadas como CORTE
        ]

        # **Contar quantas vezes cada Nome Vulgar aparece por UT**
        df_contagem = df_filtrado.groupby(["UT", "Nome Vulgar"], as_index=False).size()

        # **Renomear a coluna de contagem**
        df_contagem.columns = ["UT", "Nome Vulgar", "Quantidade"]

        # **Fazer merge para trazer Situação e UT_AREA_HA**
        df_contagem = df_contagem.merge(
            df_saida[["UT", "Nome Vulgar", "Situacao", "UT_AREA_HA"]].drop_duplicates(), 
            on=["UT", "Nome Vulgar"], 
            how="left"
        )
        
        # **Definir a função de substituição**
        def definir_substituta_vuneravel(quantidade, Situacao, area_hect):
            if Situacao == "SEM RESTRIÇÃO":
                x = np.ceil(quantidade * 0.1)
                y = np.ceil(area_hect * 0.03)
                return max(x, y)

            if Situacao == "VULNERÁVEL":
                x = np.ceil(quantidade * 0.15)
                y = np.ceil(area_hect * 0.04)
                return max(x, y)
            
            return 0  # Se não for vulnerável nem sem restrição, retorna 0

        # **Aplicar a função para calcular o valor de substituição**
        df_contagem["Valor_Substituta"] = df_contagem.apply(
            lambda row: definir_substituta_vuneravel(row["Quantidade"], row["Situacao"], row["UT_AREA_HA"]), 
            axis=1
        )

        # **Exibir os resultados finais**
        #print(df_contagem)

        ##### Classificar Substitutas 
        # Filtrar apenas as árvores que estão como CORTE e com Nome Vulgar nos selecionados
        df_filtrado = df_saida[
            (df_saida["Categoria"] == "CORTE") & 
            (df_saida["Nome Vulgar"].isin(nomes_selecionados))
        ].copy()
        
        # ordenação e prioridade para substituta 
        ordering_mode, df_filtrado
        if ordering_mode == "QF > Vol":
            df_filtrado.sort_values(by=["UT", "QF", "Vol"], ascending=[True, False, True], inplace=True)
            print("-----QF > Vol-----")
        elif ordering_mode == "Vol > QF":
            df_filtrado.sort_values(by=["UT", "Vol", "QF"], ascending=[True, True, False], inplace=True)
            print("----------")
        elif ordering_mode == "Apenas QF":
            df_filtrado.sort_values(by=["UT", "QF"], ascending=[True, False], inplace=True)
            print("----------")
        elif ordering_mode == "Apenas Vol":
            df_filtrado.sort_values(by=["UT", "Vol"], ascending=[True, True], inplace=True)
            print("----------")
        else:
            print("Modo de ordenação não reconhecido.")
            return

        

        # Garantir que df_contagem tenha apenas uma linha por UT e Nome Vulgar
        df_contagem_agg = df_contagem.groupby(["UT", "Nome Vulgar"], as_index=False).agg({"Valor_Substituta": "sum"})

        # Mesclar com df_contagem_agg para garantir que a quantidade de substitutas seja específica para cada UT e Nome Vulgar
        df_filtrado = df_filtrado.merge(df_contagem_agg, on=["UT", "Nome Vulgar"], how="left")

        print("\n--- Validação: df_filtrado após o merge ---")
        print(df_filtrado.head())

        # Função para definir as árvores substitutas corretamente
        def definir_substituta(df):
            df["Marcador"] = False  # Criar coluna auxiliar para marcar substitutas

            # Iterar por UT e Nome Vulgar
            for (ut, nome), grupo in df.groupby(["UT", "Nome Vulgar"]):
                quantidade_substituir = grupo["Valor_Substituta"].iloc[0]  # Obter a quantidade correta

                # Garantir que não substituímos mais do que o disponível no grupo
                if pd.notna(quantidade_substituir) and quantidade_substituir > 0:
                    quantidade_substituir = min(int(quantidade_substituir), len(grupo))
                    indices_para_substituir = grupo.index[:quantidade_substituir]
                    df.loc[indices_para_substituir, "Marcador"] = True

            # Aplicar a substituição apenas para os marcados
            df.loc[df["Marcador"], "Categoria"] = "SUBSTITUTA"
            df.drop(columns=["Marcador"], inplace=True)

            return df

        # Aplicar a função para categorizar corretamente como SUBSTITUTA
        df_filtrado = definir_substituta(df_filtrado)

        # Filtrar apenas os registros que realmente foram substituídos
        df_substituta = df_filtrado[df_filtrado["Categoria"] == "SUBSTITUTA"][["UT", "Nome Vulgar", "Categoria"]]

        print("\n--- df_substituta (Apenas os registros que devem ser SUBSTITUTA) ---")
        print(df_substituta.drop_duplicates())

        # Atualizar SOMENTE os registros corretos em df_saida
        # Evitar conflitos mantendo índices consistentes
        df_saida.set_index(["UT", "Nome Vulgar", "Faixa", "Placa"], inplace=True)
        df_filtrado.set_index(["UT", "Nome Vulgar", "Faixa", "Placa"], inplace=True)

        # Somente substituir onde há correspondência exata
        df_saida.update(df_filtrado["Categoria"])

        # Resetar índice após atualização
        df_saida.reset_index(inplace=True)

        ####
        # **Contar a quantidade de árvores "CORTE" por UT e Nome Vulgar**
        df_contagem_corte = df_saida[df_saida["Categoria"] == "CORTE"].groupby(["UT", "Nome Vulgar"], as_index=False).size()
        df_contagem_corte.rename(columns={"size": "Qtd_Corte"}, inplace=True)

        # **Mesclar com a contagem de cortes para verificar onde não há mais cortes**
        df_verificacao = df_substituta.merge(df_contagem_corte, on=["UT", "Nome Vulgar"], how="left").fillna(0)

        # **Criar um identificador para marcar onde NÃO EXISTE "CORTE"**
        df_verificacao["Marcar_REM"] = df_verificacao["Qtd_Corte"] == 0

        # **Criar uma lista de tuplas (UT, Nome Vulgar) onde as substitutas precisam virar REMANESCENTE**
        remover_tuplas = df_verificacao[df_verificacao["Marcar_REM"]][["UT", "Nome Vulgar"]].apply(tuple, axis=1).tolist()

        # **Atualizar df_saida para transformar "SUBSTITUTA" em "REMANESCENTE" onde não há cortes**
        df_saida["Categoria"] = df_saida.apply(
            lambda row: "REMANESCENTE" if (row["UT"], row["Nome Vulgar"]) in remover_tuplas and row["Categoria"] == "SUBSTITUTA" else row["Categoria"],
            axis=1
        )

        # **Verificar os resultados corrigidos**
        print("\n--- Linhas que viraram REMANESCENTE porque não há mais CORTE dentro da UT ---")
        print(df_saida[df_saida["Categoria"] == "REMANESCENTE"][["UT", "Nome Vulgar", "Categoria"]].drop_duplicates())

        print("\n--- Contagem Final por Categoria ---")
        print(df_saida["Categoria"].value_counts())


        contagem_categorias = df_saida["Categoria"].value_counts()

        print("Contagem por Categoria:")
        print(f"CORTE: {contagem_categorias.get('CORTE', 0)}")
        print(f"SUBSTITUTA: {contagem_categorias.get('SUBSTITUTA', 0)}")
        print(f"REMANESCENTE: {contagem_categorias.get('REMANESCENTE', 0)}")

        print(f"Numero total de linhas em df_saida: {len(df_saida)}")
        tabelaDeResumo()
        finalProcesso = time.time()
        print(f"Processamento realizado em {finalProcesso - inicioProcesso:.2f} s")
        
        inicioTimeSalvar = time.time()
        #organizando as colunas
        df_saida = df_saida[COLUNAS_SAIDA]

        if contagem_categorias.get("CORTE", 0) > 0:
            print(f"CORTE: {contagem_categorias['CORTE']}")
            print(f"SUBSTITUTA: {contagem_categorias.get('SUBSTITUTA', 0)}")
            print(f"REMANESCENTE: {contagem_categorias.get('REMANESCENTE', 0)}")

            df_saida = ajustarVolumeHect()
            
            print("dentro de processo")
            print(df_saida)
            if save == True :
                # Exporta para Excel usando o df_modificado com os cálculos originais
                df_export = df_saida[["UT", "Faixa", "Placa", "Nome Vulgar", "CAP", "CAP_a", "mCAP", "H", "H_a", "mH", "QF",
                                        "X", "Y", "DAP", "DAP_a", "Vol", "Vol_a","Lat", "Long", "DM", "OBS",
                                        "Categoria", "Situacao"]]
                diretorio = os.path.dirname(entrada1_var.get())
                arquivo_saida = os.path.join(diretorio, "Planilha Processada - IFDIGITAL 3.0.xlsx")
                df_export.to_excel(arquivo_saida, index=False, engine="xlsxwriter")
                finalTimeSalvar = time.time()

                # finalizando barra de progresso
                # Para e esconde a barra de progresso e exibe novamente o botão
                progress_bar.stop()
                progress_bar.pack_forget()
                button_process.pack(pady=0)
                
                tk.messagebox.showinfo(
                "SUCESSO ! ",
                f"Arquivo salvo em {diretorio} \nProcessamento realizado em {finalProcesso - inicioProcesso:.2f} segundos e o arquivo salvo em {finalTimeSalvar - inicioTimeSalvar:.2f} s ."
                )
            else :
                print("fim de processamento")

            #salvar o arquivo no diretório
            # diretorio = os.path.dirname(entrada1_var.get())
            # arquivo_saida = os.path.join(diretorio, "Planilha Processada - IFDIGITAL 3.0.xlsx")
            # df_saida.to_excel(arquivo_saida,index=False,engine="xlsxwriter")

            

            # finalizando barra de progresso
            # Para e esconde a barra de progresso e exibe novamente o botão
            progress_bar.stop()
            progress_bar.pack_forget()
            button_process.pack(pady=0)
            
        else:
            progress_bar.stop()
            progress_bar.pack_forget()
            button_process.pack(pady=0)

            tk.messagebox.showinfo(
            "ERRO !",
            "Nem uma Espécie foi categorizada como corte"
            )

    except Exception as e:
        # finalizando barra de progresso
        # Para e esconde a barra de progresso e exibe novamente o botão
        progress_bar.stop()
        progress_bar.pack_forget()
        button_process.pack(pady=0)
        tk.messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")

def iniciar_processamento(save_planilha):
    config = configparser.ConfigParser()
    config.read('config.ini')
    """Inicia o processamento em uma thread separada."""
    thread = threading.Thread(
    target=processar_planilhas,
    args=(save_planilha,)
    )
    thread.daemon = True  # Fecha a thread quando a interface é fechada
    thread.start()

# Função para abrir a janela de valores padrões ao clicar no botão
def abrir_janela_valores_padroes_callback():
    # Abre a janela de valores padrões e bloqueia a janela principal
    janela_padrao = abrir_janela_valores_padroes(app)
    
    # Aguarda até que a janela secundária seja fechada
    app.wait_window(janela_padrao)

# Índices das colunas que podem ser editadas
colunas_editaveis = [ "CAP", "H"]

def editar_celula_volume(event):

    def editarEspeciesUT(event):
        global dados_editados_por_ut
        item = table_ut_vol.selection()[0]
        ut = table_ut_vol.item(item, "values")[0]
        valores_ut = table_ut_vol.item(item, "values")
        hectares = float(valores_ut[1])

        df_filtrado = df_saida[
            (df_saida["UT"] == int(ut)) &
            (df_saida["Categoria"].str.upper().str.strip() == "CORTE")
        ]

        agrupado = df_filtrado.groupby("Nome Vulgar").agg(
            num_arvores=('Nome Vulgar', 'count'),
            volume_total=('Vol_a', 'sum'),
        ).reset_index()
        agrupado["vol_por_ha"] = agrupado["volume_total"] / hectares

        nova_janela = tk.Toplevel()
        nova_janela.title(f"Espécies da UT {ut}")
        nova_janela.geometry("1050x700")


        tabela = ttk.Treeview(nova_janela, columns=("Nome", "n° Árvores", "Volume Total", "Vol/ha", "DAP <", "QF >=", "F_REM"),
                            show="headings", height=20, style="verde.Treeview")

        colunas = ("Nome", "n° Árvores", "Volume Total", "Vol/ha", "DAP <", "QF >=", "F_REM")
        for col in colunas:
            tabela.heading(col, text=col)
            tabela.column(col, width=140, anchor="center")

        tabela.pack(pady=10, padx=10, fill="x")

        # Preenche a tabela com cores alternadas
        for i, row in agrupado.iterrows():
            nome = row["Nome Vulgar"]
            dados_salvos = dados_editados_por_ut.get(ut, {}).get(nome, {})

            valores = [
                nome,
                row["num_arvores"],
                f"{row['volume_total']:.2f}",
                f"{row['vol_por_ha']:.2f}",
                dados_salvos.get("DAP <", ""),
                dados_salvos.get("QF >=", ""),
                dados_salvos.get("F_REM", "NÃO")
            ]

            tag = 'verde' if i % 2 == 0 else 'branco'
            tabela.insert("", "end", iid=nome, values=valores, tags=(tag,))

        tabela.tag_configure('verde', background="#e5fbe0")
        tabela.tag_configure('branco', background="#ffffff")

        # Área de edição
        frame_edicao = ttk.LabelFrame(nova_janela, text="Editar Espécie Selecionada")
        frame_edicao.pack(fill="x", padx=10, pady=10)

        entradas = {}
        for i, campo in enumerate(colunas[4:]):
            ttk.Label(frame_edicao, text=campo).grid(row=0, column=i, padx=5)
            if campo == "F_REM":
                combo = ttk.Combobox(frame_edicao, values=["SIM","NÃO"], state="readonly", width=10)
                combo.grid(row=1, column=i, padx=5)
                entradas[campo] = combo
            else:
                entry = ttk.Entry(frame_edicao, width=10)
                entry.grid(row=1, column=i, padx=5)
                entradas[campo] = entry

        especie_selecionada = tk.StringVar()

        def ao_selecionar_linha(event):
            item = tabela.selection()
            if item:
                nome = item[0]
                especie_selecionada.set(nome)
                dados = tabela.item(nome, "values")
                for i, campo in enumerate(colunas[4:]):
                    entradas[campo].delete(0, tk.END)
                    entradas[campo].insert(0, dados[i + 4])

        tabela.bind("<<TreeviewSelect>>", ao_selecionar_linha)

        def salvar_dados():
            nome = especie_selecionada.get()
            if not nome:
                return
            if ut not in dados_editados_por_ut:
                dados_editados_por_ut[ut] = {}
            dados_editados_por_ut[ut][nome] = {
                campo: entradas[campo].get() for campo in entradas
            }
            print(f"Salvo UT {ut} - {nome}: {dados_editados_por_ut[ut][nome]}")
            # Atualiza visualmente a tabela
            valores_atualizados = tabela.item(nome, "values")[:4] + tuple(entradas[campo].get() for campo in entradas)
            tabela.item(nome, values=valores_atualizados)

        def excluir_alteracoes_atuais():
            if ut in dados_editados_por_ut:
                print("Antes da exclusão:")
                print(dados_editados_por_ut)

                del dados_editados_por_ut[ut]

                print("Depois da exclusão:")
                print(dados_editados_por_ut)
                nova_janela.destroy()
                iniciar_processamento(False)
                messagebox.showinfo("Alterações Excluídas", f"Todas as alterações da UT '{ut}' foram removidas.")
            else:
                messagebox.showinfo("Nenhuma Alteração", f"Não há alterações registradas para a UT '{ut}'.")
            


        # Frame para os botões lado a lado
        frame_botoes = ttk.Frame(nova_janela)
        frame_botoes.pack(pady=10)

        # Botão Salvar
        btn_salvar = ttk.Button(frame_botoes, text="Salvar Alterações", command=salvar_dados)
        btn_salvar.grid(row=0, column=0, padx=5)

        # Botão Aplicar
        btn_aplicar = ttk.Button(frame_botoes, text="Aplicar", command=lambda: iniciar_processamento(False))
        btn_aplicar.grid(row=0, column=1, padx=5)

        btn_excluir = ttk.Button(frame_botoes, text="Excluir Alterações", command=excluir_alteracoes_atuais)
        btn_excluir.grid(row=0, column=2, padx=5)




    def ajustarCAP_ALT(item, event):
        # Agora que temos o item e o evento, podemos usar o evento para identificar a coluna
        coluna_selecionada = table_ut_vol.identify_column(event.x)
        print(f"Coluna clicada: {coluna_selecionada}")
        # Lógica de ajuste de CAP/H
        print(f"Item selecionado: {item}")
        # Aqui você pode adicionar o que precisa fazer com o item selecionado
        """
        Permite editar as colunas CAP e H ao dar duplo clique em uma célula da Treeview (table_ut_vol).
        Apenas colunas especificadas em 'colunas_editaveis' podem ser editadas.
        """
        global df_tabelaDeAjusteVol, df_valores_atualizados

        # Captura o item (linha) selecionado
        item_selecionado = table_ut_vol.focus()
        if not item_selecionado:
            return

        # Identifica a coluna clicada a partir da coordenada x do evento
        coluna_selecionada = table_ut_vol.identify_column(event.x)
        try:
            # Converte, por exemplo, "#2" para índice 1 (0-indexado)
            col_index = int(coluna_selecionada.lstrip("#")) - 1
        except Exception as e:
            print("Erro ao identificar a coluna:", e)
            return

        # Obtém o nome da coluna correspondente
        col_nome = table_ut_vol["columns"][col_index]
        # Verifica se a coluna é editável (lista global definida, por exemplo: colunas_editaveis = ["CAP", "H"])
        if col_nome not in colunas_editaveis:
            return

        # Obtém as coordenadas da célula para posicionar o widget de edição (Entry)
        bbox = table_ut_vol.bbox(item_selecionado, col_index)
        if not bbox:
            return
        x, y, largura, altura = bbox

        # Obtém o valor atual da célula
        valores = list(table_ut_vol.item(item_selecionado, "values"))
        valor_atual = valores[col_index]

        # Cria o widget Entry para edição, posicionando-o na célula
        entry = tk.Entry(table_ut_vol)
        entry.place(x=x, y=y, width=largura, height=altura)
        entry.insert(0, valor_atual)
        entry.focus()

        def salvar_novo_valor(event=None):
            """
            Função interna para salvar o novo valor digitado.
            Atualiza a Treeview e os DataFrames df_tabelaDeAjusteVol e df_valores_atualizados.
            """
            global df_tabelaDeAjusteVol, df_valores_atualizados

            novo_valor_str = entry.get()
            try:
                novo_valor = float(novo_valor_str)
            except ValueError:
                entry.destroy()
                return

            # Atualiza a lista de valores da linha na Treeview
            valores[col_index] = novo_valor
            table_ut_vol.item(item_selecionado, values=valores)

            # Obtém o valor da primeira coluna (UT) e converte para float
            try:
                ut_val = float(valores[0])
            except ValueError:
                print("Erro ao converter UT para float.")
                entry.destroy()
                return

            # Verifica se "ut" existe em df_tabelaDeAjusteVol
            if "ut" not in df_tabelaDeAjusteVol.columns:
                print("Erro: A coluna 'ut' não existe no DataFrame!")
                entry.destroy()
                return

            # Localiza a linha no DataFrame correspondente à UT
            index_df_series = df_tabelaDeAjusteVol[df_tabelaDeAjusteVol["ut"] == ut_val].index
            if index_df_series.empty:
                print(f"Erro: UT {ut_val} não encontrado no DataFrame!")
                entry.destroy()
                return
            index_df = index_df_series[0]

            # Atualiza o valor da coluna editada no DataFrame de ajuste
            df_tabelaDeAjusteVol.at[index_df, col_nome] = novo_valor

            # Se a coluna editada for CAP ou H, atualiza também o DataFrame global de valores atualizados
            if col_nome in ["CAP", "H"]:
                # Se df_valores_atualizados não existir ou estiver vazio, inicializa-o
                if 'df_valores_atualizados' not in globals() or df_valores_atualizados.empty:
                    df_valores_atualizados = pd.DataFrame(columns=["ut", "CAP", "H"])
                # Procura se já existe uma linha para essa UT
                idx_val_series = df_valores_atualizados[df_valores_atualizados["ut"] == ut_val].index
                if not idx_val_series.empty:
                    df_valores_atualizados.at[idx_val_series[0], col_nome] = novo_valor
                else:
                    nova_linha = {"ut": ut_val, "CAP": np.nan, "H": np.nan}
                    nova_linha[col_nome] = novo_valor
                    df_valores_atualizados = pd.concat([df_valores_atualizados, pd.DataFrame([nova_linha])], ignore_index=True)

            print(f"Valor atualizado: UT={ut_val}, {col_nome}={novo_valor}")
            print("DataFrame de ajuste atualizado:")
            print(df_tabelaDeAjusteVol)
            print("DataFrame de valores atualizados:")
            print(df_valores_atualizados)

            ajustarVolumeHect()

            entry.destroy()
        
        # Vincula a função salvar_novo_valor aos eventos Return e FocusOut
        entry.bind("<Return>", salvar_novo_valor)
        entry.bind("<FocusOut>", salvar_novo_valor)

     # Pega o item selecionado
    item = table_ut_vol.selection()[0]
    
    # Identifica a coluna clicada (usando o evento)
    coluna_clicada = table_ut_vol.identify_column(event.x)  # Identifica a coluna pelo eixo X do evento
    
    # Verifica a coluna clicada e chama a função correspondente
    if coluna_clicada == "#1":  # A coluna UT é a primeira (coluna #1)
        # Se for a coluna "UT", chama a função para editar a célula de UT
        editarEspeciesUT(item)
    elif coluna_clicada == "#10":  # A coluna CAP é a coluna #10
        # Se for a coluna "CAP", chama a função para editar a célula de CAP
        ajustarCAP_ALT(item,event)
    elif coluna_clicada == "#11":  # A coluna H é a coluna #11
        # Se for a coluna "H", chama a função para editar a célula de H
        ajustarCAP_ALT(item,event)
    else:
        print(f"Cliquei na coluna {coluna_clicada}, mas não tem ação associada.")

def alternar_tabela():
    global tabela_visivel

    if tabela_visivel:
        frame_tabela2.pack_forget()
        frame_resumo_especie.pack(fill="both", expand=True)
        botao_trocar_tabela.config(text="Mostrar Tabela de Edição")
        tabela_visivel = False
    else:
        frame_resumo_especie.pack_forget()
        frame_tabela2.pack(fill="both", expand=True)
        botao_trocar_tabela.config(text="Ocultar Tabela de Edição")
        tabela_visivel = True


def definir_icone(app):
    try:
        # Verifica se o script está sendo executado como executável (PyInstaller)
        if getattr(sys, 'frozen', False):
            # Quando o aplicativo é compilado, os arquivos são extraídos para o diretório _MEIPASS
            base_path = sys._MEIPASS
        else:
            # Durante o desenvolvimento, usa o diretório do script
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        # Caminho do ícone relativo ao diretório correto
        icone_path = os.path.join(base_path, "", "icone ifdigital.ico")
        
        # Define o ícone para a janela
        app.iconbitmap(icone_path)
        print(f"Ícone carregado com sucesso a partir de: {icone_path}")

    except Exception as e:
        print(f"Erro ao carregar o ícone: {e}")

# Criação da janela principal
app = tk.Tk()
app.title("IFDIGITAL 3.0")

# Chama a função para definir o ícone
definir_icone(app)

# Dimensões da janela
largura_janela = 1600
altura_janela = 900

# Obter largura e altura da tela
largura_tela = app.winfo_screenwidth()
altura_tela = app.winfo_screenheight()

# Estilo para alternar as cores das linhas
style = ttk.Style()
style.configure("Treeview", rowheight=20)
style.map("Treeview", background=[('selected', '#38761d')], foreground=[('selected', 'white')])
style.configure("verde.Treeview", background="white")
style.layout("verde.Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

# Calcular coordenadas para centralizar a janela
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

# Definir a geometria da janela com posição centralizada
app.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

# Impedir a maximização da janela
app.resizable(False, False)  # Permite redimensionamento
app.maxsize(largura_janela, altura_janela)  # Impede que a janela seja maximizada

# Carregar imagem de fundo corretamente
caminho_imagem = os.path.join("src", "01florest.png")  # Ajuste conforme necessário

if not os.path.exists(caminho_imagem):
    print(f"Erro: Arquivo {caminho_imagem} não encontrado!")

# Variável global para manter a imagem na memória
global fundo_tk  
imagem_fundo = Image.open(caminho_imagem)
imagem_fundo = imagem_fundo.resize((largura_janela, altura_janela), Image.Resampling.LANCZOS)
fundo_tk = ImageTk.PhotoImage(imagem_fundo)

# Criar um canvas e definir a imagem de fundo
canvas = tk.Canvas(app, width=largura_janela, height=altura_janela)
canvas.pack(fill="both", expand=True)  # Faz o canvas ocupar toda a janela
canvas.create_image(0, 0, image=fundo_tk, anchor="nw")  # Define a imagem de fundo

# Variáveis globais
entrada1_var = tk.StringVar()
entrada2_var = tk.StringVar()
pesquisa_var = tk.StringVar()

# Frame para entrada de arquivos
frame_inputs = ttk.LabelFrame(app, text="Entrada de Arquivos", padding=(10, 10))
frame_inputs.place(x=10, y=10)  # Posiciona o frame sobre o canvas

ttk.Label(frame_inputs, text="INVENTÁRIO :").grid(row=0, column=0, sticky="w")
ttk.Entry(frame_inputs, textvariable=entrada1_var, width=60).grid(row=0, column=1, pady=5, padx=5)
ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("principal")).grid(row=0, column=2, padx=5)

ttk.Label(frame_inputs, text="LISTA DE ESPECIES :").grid(row=1, column=0, sticky="w")
ttk.Entry(frame_inputs, textvariable=entrada2_var, width=60).grid(row=1, column=1, pady=5, padx=5)
ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("secundaria")).grid(row=1, column=2, padx=5)

# Frame para organizar os widgets lado a lado
frame_lado_a_lado = tk.Frame(app)
frame_lado_a_lado.place(x=30, y=125)

# Botão para modificar a filtragem para REM
botao_modificar_filtro = ttk.Button(
    frame_lado_a_lado,
    text="Modificar Filtragem para REM",
    command=abrir_janela_valores_padroes_callback
)
botao_modificar_filtro.pack(side=tk.LEFT, padx=50)

# Label de texto antes do combobox
label_texto = ttk.Label(frame_lado_a_lado, text="Ordenação de Substituta:")
label_texto.pack(side=tk.LEFT, padx=(0, 0))

# Combobox com opções de ordenação
combobox = ttk.Combobox(
    frame_lado_a_lado,
    state="readonly",
    style="TCombobox",
    values=[
        "QF > Vol",
        "Vol > QF",
        "Apenas QF",
        "Apenas Vol"
    ]
)
combobox.current(0)
combobox.bind("<<ComboboxSelected>>", update_ordering_mode)
combobox.pack(side=tk.LEFT, padx=10)

# Novo botão para alternar entre a tabela de ajustes e a de resumo
botao_trocar_tabela = ttk.Button(
    frame_lado_a_lado,
    text="Mostrar Tabela de Resumo",  # ou "Mostrar Tabela de Ajustes" dependendo do estado atual
    command=alternar_tabela
)
botao_trocar_tabela.pack(side=tk.RIGHT, padx=(0))


# Exibindo o Combobox
combobox.pack(side=tk.LEFT)

# Criação dos widgets que serão atualizados
status_label = ttk.Label(app, text="")  # Inicialmente vazio
status_label.pack(pady=10)

frame_listbox_e_tabela =tk.Frame(app)
frame_listbox_e_tabela.pack(side="left", padx=10, pady=10, fill="y")

# Frame para Listboxes

frame_listbox = ttk.LabelFrame(frame_listbox_e_tabela, text="Seleção de Nomes Vulgares", padding=(10, 10))
frame_listbox.pack(side="left", padx=10, pady=0, fill="y")


# Criando um Frame para alinhar a Label e a Entry horizontalmente
frame_pesquisa = tk.Frame(frame_listbox)
frame_pesquisa.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")

# Adicionando a Label e a Entry dentro do frame_pesquisa
ttk.Label(frame_pesquisa, text="Pesquisar:").pack(side="left")
pesquisa_entry = ttk.Entry(frame_pesquisa, textvariable=pesquisa_var, width=40)
pesquisa_entry.pack(side="left", padx=5)

# Vinculando evento de pesquisa
pesquisa_entry.bind("<KeyRelease>", pesquisar_nomes)

colunas_selecionados = ("Nome", "DAP <", "DAP >=", "QF = ", "H >")

table_selecionados = ttk.Treeview(frame_listbox, columns=colunas_selecionados, show="headings", height=15)
for col in colunas_selecionados:
    table_selecionados.heading(col, text=col)
    table_selecionados.column(col, width=50, anchor="center")
    table_selecionados.column("Nome", width=200, anchor="w") 
table_selecionados.grid(row=1, column=1, padx=10, pady=10)


# Criando a Listbox com os nomes vulgares
listbox_nomes_vulgares = tk.Listbox(frame_listbox, selectmode=tk.SINGLE, width=25, height=20)
listbox_nomes_vulgares.grid(row=1, column=0, padx=10, pady=10)

# Vinculando o clique na Listbox para adicionar à tabela
listbox_nomes_vulgares.bind("<ButtonRelease-1>", adicionar_selecao)

# Vinculando o duplo clique na tabela para edição
table_selecionados.bind("<Double-1>", editar_linha)

# Criar um frame para os botões
frame_botoes = tk.Frame(frame_listbox)
frame_botoes.grid(row=6, column=0, columnspan=3, pady=0)  # Ocupa 3 colunas

# Criar os botões dentro do frame_botoes
btn_todos = ttk.Button(frame_botoes, text="Selecionar Todos", command=selecionar_todos)
btn_todos.pack(side="left", padx=10)

btn_remover = ttk.Button(frame_botoes, text="Remover Último", command=remover_ultimo_selecionado)
btn_remover.pack(side="left", padx=10)

btn_limpar = ttk.Button(frame_botoes, text="Limpar Lista", command=limpar_lista_selecionados)
btn_limpar.pack(side="left", padx=10)

#frma para a tabela de ajuste de volume por hectar 

frame_tabela_container = ttk.Frame(frame_listbox_e_tabela)
frame_tabela_container.pack(side="right", fill="both", expand=True)

# Frame da tabela de edição
frame_tabela2 = ttk.Frame(frame_tabela_container)

colunas_tabela2 = ("UT", "Hectares", "n° Árv", "Vol", "Vol_Max", "Diminuir", "Aumentar", "Vol/ha",
                   "DAP %", "CAP", "H")

table_ut_vol = ttk.Treeview(frame_tabela2, columns=colunas_tabela2, show="headings", height=22)
for col in colunas_tabela2:
    table_ut_vol.heading(col, text=col)
    table_ut_vol.column(col, width=75, anchor="center")
table_ut_vol.pack(fill="both", expand=True)
table_ut_vol.bind("<Double-1>", editar_celula_volume)

frame_tabela2.pack(fill="both", expand=True)  # Mostra a de edição inicialmente
tabela_visivel = True

# Frame da tabela de resumo
frame_resumo_especie = ttk.Frame(frame_tabela_container)

colunas_resumo_especie = (
    "Espécie", "n° Árvores", "Vol/Árvore", "Vol Ajustado",
    "Vol/Árvore Ajustado", "Vol/ha", "CAP_min"
)

table_resumo_especie = ttk.Treeview(
    frame_resumo_especie,
    columns=colunas_resumo_especie,
    show="headings",
    height=22
)
for col in colunas_resumo_especie:
    table_resumo_especie.heading(col, text=col)
    table_resumo_especie.column(col, width=120, anchor="center")

# Scroll
scroll_resumo_especie = ttk.Scrollbar(frame_resumo_especie, orient="vertical", command=table_resumo_especie.yview)
table_resumo_especie.configure(yscrollcommand=scroll_resumo_especie.set)

# Packing
table_resumo_especie.pack(side="left", fill="both", expand=True)
scroll_resumo_especie.pack(side="right", fill="y")

# Frame para o botão de processamento

frame_secundario = ttk.Frame(app)
frame_secundario.place(x= 30, y=125)
button_process = ttk.Button(frame_secundario, text="Processar Planilhas",  command=lambda: iniciar_processamento(True), width=40)
button_process.pack(padx=(0,0), pady=(0,0))

# Barra de progresso (inicialmente não exibida)
progress_bar = ttk.Progressbar(frame_secundario, mode='indeterminate', length=300)

frame_listbox_e_tabela.place_forget()
frame_secundario.place_forget()  # Inicialmente escondido

carregar_planilha_salva("principal")
carregar_planilha_salva("secundaria")

app.mainloop()
