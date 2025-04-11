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
import math


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
# Inicializando o df_resumo globalmente como DataFrame vazio
df_resumo = pd.DataFrame(columns=[
    "CAPmin", "Espécie", "n° Árvores", "Vol", "Vol/Árvore", 
    "Volume_a", "Vol_a/Árvore", "Vol/Hect"
])
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
    # Caminho para o ícone da janela
    icone_path = resource_path("src/img/icoGreenFlorest.ico")

    # Define o ícone da aplicação
    popup.iconbitmap(icone_path)

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
    # Recarrega os valores do config.ini
    config.read("config.ini")

    dap_max = config.getfloat("DEFAULT", "dapmax", fallback=0.5)
    dap_min = config.getfloat("DEFAULT", "dapmin", fallback=5)
    qf = config.getint("DEFAULT", "qf", fallback=3)
    alt = config.getfloat("DEFAULT", "alt", fallback=0)
    cap = config.getfloat("DEFAULT", "cap", fallback=2.5)

    # Se alt for 0, substituir por string vazia
    alt = "" if alt == 0 else alt

    count_existentes = len(table_selecionados.get_children())

    for i, nome in enumerate(nomes_vulgares):
        # Verifica se já está na tabela
        is_duplicate = any(table_selecionados.item(child, "values")[0] == nome for child in table_selecionados.get_children())
        
        if not is_duplicate:
            # Alterna as cores
            tag = 'verde_claro' if (count_existentes + i) % 2 == 0 else 'branca'

            valores_atualizados = (nome, dap_max, dap_min, qf, alt, cap)
            table_selecionados.insert("", "end", values=valores_atualizados, tags=(tag,))
            print(f"{nome} foi adicionado à tabela.")
        else:
            print(f"{nome} já está na tabela.")

    table_selecionados.tag_configure('branca', background='white')
    table_selecionados.tag_configure('verde_claro', background='#d3f8e2')

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
    global df_resumo

    # Limpa o DataFrame de resumo antes de preencher novamente
    df_resumo = pd.DataFrame(columns=df_resumo.columns)

    # Limpa a tabela antes de preencher de novo
    for item in table_resumo_especie.get_children():
        table_resumo_especie.delete(item)

    # Cópia e filtro
    df_resumo_especies = df_saida.copy()
    df_resumo_especies = df_resumo_especies[df_resumo_especies["Categoria"] == "CORTE"]

    # Agrupamento por espécie
    agrupado = df_resumo_especies.groupby("Nome Vulgar")

    for i, (especie, grupo) in enumerate(agrupado):
        qtd_arvores = len(grupo)
        soma_vol = grupo["Vol"].sum()
        soma_vol_a = grupo["Vol_a"].sum()
        vol_por_arvore = soma_vol / qtd_arvores if qtd_arvores > 0 else 0.0
        vol_a_por_arvore = soma_vol_a / qtd_arvores if qtd_arvores > 0 else 0.0

        df_uts_hectares = grupo[["UT", "UT_AREA_HA"]].drop_duplicates()
        soma_hectares = df_uts_hectares["UT_AREA_HA"].sum()
        vol_ha = soma_vol_a / soma_hectares if soma_hectares > 0 else 0.0

        cap_min = grupo["CAP_a"].min()

        # Cria uma tupla com os valores calculados
        valores = (
            round(cap_min, 2),
            especie,
            qtd_arvores,
            round(soma_vol, 2),
            round(vol_por_arvore, 2),
            round(soma_vol_a, 2),
            round(vol_a_por_arvore, 2),
            round(vol_ha, 2)
        )

        # Alterna entre tags
        tag = "linha_verde" if i % 2 == 0 else "linha_branca"
        table_resumo_especie.insert("", "end", values=valores, tags=(tag,))

        # Adiciona os valores à lista de dados_resumo
        df_resumo = pd.concat([df_resumo, pd.DataFrame([valores], columns=df_resumo.columns)], ignore_index=True)
def definir_e_Recuperar_mCAP_e_mH():
    global df_tabelaDeAjusteVol, df_valores_atualizados,df_saida, dados_editados_por_ut
    # Cria uma cópia do DataFrame original para não modificá-lo
    df_modificado = df_saida.copy()

    # Cria as colunas mCAP e mH com valores padrão: mCAP = 1, mH = 0
    df_modificado["mCAP"] = 1
    df_modificado["mH"] = 0

    # Verifica se o DataFrame de valores atualizados existe e não está vazio
    if 'df_valores_atualizados' in globals() and not df_valores_atualizados.empty:
        for col in ["mCAP", "mH"]:
            # Buscar os valores originais com base no novo nome
            nome_origem = col[1:]  # remove o 'm' do início
            df_valores_atualizados[col] = df_valores_atualizados.get(nome_origem, df_valores_atualizados.get(col))

        # Realiza o merge entre os DataFrames df_modificado e df_valores_atualizados
        df_modificado = df_modificado.merge(
            df_valores_atualizados[["ut", "mCAP", "mH"]],
            left_on="UT", right_on="ut", how="left", suffixes=("", "_novo")
        )

        # Atualiza os valores de mCAP e mH com os valores de df_valores_atualizados, caso existam
        for col in ["mCAP", "mH"]:
            if f"{col}_novo" in df_modificado.columns:
                novo_val = df_modificado[f"{col}_novo"]
                df_modificado[col] = np.where(novo_val.notna(), novo_val, df_modificado[col])

        # Remove as colunas temporárias criadas no merge
        cols_to_drop = [f"{col}_novo" for col in ["mCAP", "mH"]]
        df_modificado.drop(columns=cols_to_drop, inplace=True)

        # Remove a coluna 'ut' se existir
        if "ut" in df_modificado.columns and "UT" in df_modificado.columns:
            df_modificado.drop(columns=["ut"], inplace=True)
    # Verifica se o dicionário 'dados_editados_por_ut' existe e não está vazio
    if 'dados_editados_por_ut' in globals() and dados_editados_por_ut:

        for idx, row in df_modificado.iterrows():
            ut = str(row["UT"]).strip()
            especie = str(row.get("Nome Vulgar", "")).strip()

            if ut in dados_editados_por_ut and especie in dados_editados_por_ut[ut]:
                cap = dados_editados_por_ut[ut][especie].get("CAP", "")
                h = dados_editados_por_ut[ut][especie].get("H", "")

                # Atribui os valores se forem válidos (não vazios)
                if cap != "":
                    df_modificado.at[idx, "mCAP"] = float(cap)
                if h != "":
                    df_modificado.at[idx, "mH"] = float(h)

def adicionarColunasAuxiliares():
    
            
    df_modificado = definir_e_Recuperar_mCAP_e_mH()        

    df_modificado["ut"] = df_modificado["UT"]
    df_modificado["Hectares"] = df_modificado["UT_AREA_HA"]

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

    return df_modificado

def adicionarDAP_a():
    
    df_modificado = definir_e_Recuperar_mCAP_e_mH()

    # Calculando a coluna 'CAP_a' somente para linhas com 'Categoria' igual a 'CORTE'
    df_modificado["CAP_a"] = df_modificado.apply(
        lambda row: round(row["CAP"] * row["mCAP"]) if row["Categoria"] == "CORTE" else row["CAP"],
        axis=1
    )

    df_modificado["DAP_a"] = df_modificado.apply(
        lambda row: ((row["CAP_a"] / np.pi) / 100) if row["Categoria"] == "CORTE" else row["DAP"],
        axis=1
    )
def ajustarVolumeHect():
    global df_tabelaDeAjusteVol, df_valores_atualizados,df_saida, dados_editados_por_ut
    

    df_modificado = adicionarColunasAuxiliares()


    df_filtrado = df_modificado[df_modificado["Categoria"].isin(["CORTE"])]

    if df_filtrado.empty:
        print("Nenhuma árvore foi categorizada como 'CORTE'.")
        df_tabelaDeAjusteVol = df_modificado[["ut", "Hectares", "mCAP", "mH", "DAP_a"]].drop_duplicates()
        df_tabelaDeAjusteVol["n_Árvores"] = 0
        df_tabelaDeAjusteVol["Vol"] = 0
        df_tabelaDeAjusteVol["diminuir"] = 0
        df_tabelaDeAjusteVol["aumentar"] = 0
        df_tabelaDeAjusteVol["media_dif_dap"] = 0
    else:
        # Agrupa para somar DAP e DAP_a por ut
        df_soma_dap = df_filtrado.groupby("ut")[["DAP", "DAP_a"]].sum().reset_index()

        # Calcula a diferença percentual entre DAP_a e DAP
        df_soma_dap["DAP %"] = ((df_soma_dap["DAP_a"] - df_soma_dap["DAP"]) / df_soma_dap["DAP"]) * 100

        # Agrupa para contar o número de árvores e somar o volume total por ut
        contagem_arvores = df_filtrado.groupby("ut").size().reset_index(name="n_Árvores")
        volume_total_por_ut = df_filtrado.groupby("ut")["Vol_a"].sum().reset_index()
        volume_total_por_ut.rename(columns={"Vol_a": "Vol"}, inplace=True)
        
        # Cria o DataFrame de ajuste com ut, Hectares, mCAP e mH
        df_tabelaDeAjusteVol = df_modificado[["ut", "Hectares", "mCAP", "mH", "DAP_a"]].drop_duplicates()
        print("Dados iniciais de ut, Hectares, mCAP, mH e DAP_a:")
        print(df_tabelaDeAjusteVol)

        # Calcula o volume máximo como (hectares * 30)
        df_tabelaDeAjusteVol["Vol_Max"] = df_tabelaDeAjusteVol["Hectares"] * 30

        # Faz merge para incorporar a contagem de árvores, o volume total e a diferença percentual
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(contagem_arvores, on="ut", how="left").fillna(0)
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(volume_total_por_ut, on="ut", how="left").fillna(0)
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(df_soma_dap[["ut", "DAP %"]], on="ut", how="left").fillna(0)

        # Calcula o volume por hectare (V³/ha)
        df_tabelaDeAjusteVol["Vol/Hect"] = df_tabelaDeAjusteVol.apply(
            lambda row: row["Vol"] / row["Hectares"] if row["Hectares"] > 0 else 0, axis=1
        )
        
        # Calcula as diferenças para "diminuir" e "aumentar"
        df_tabelaDeAjusteVol["diminuir"] = df_tabelaDeAjusteVol.apply(
            lambda row: row["Vol"] - row["Vol_Max"] if row["Vol"] > row["Vol_Max"] else 0, axis=1
        )
        df_tabelaDeAjusteVol["aumentar"] = df_tabelaDeAjusteVol.apply(
            lambda row: row["Vol_Max"] - row["Vol"] if row["Vol"] < row["Vol_Max"] else 0, axis=1
        )

        # Remover duplicatas por ut após os merges
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.drop_duplicates(subset=["ut"])

    df_tabelaDeAjusteVol["n_Árvores"] = df_tabelaDeAjusteVol["n_Árvores"].astype(int)
    # Imprime os nomes das colunas de df_tabelaDeAjusteVol
    print(df_tabelaDeAjusteVol.columns.tolist())

    for col in df_tabelaDeAjusteVol.columns:
        print(col)
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
        hectares_val = f"{row['Hectares']:.5f}"
        n_Árvores = f"{row['n_Árvores']:.0f}"
        Vol = f"{row['Vol']:.5f}"
        Vol_Max = f"{row['Vol_Max']:.5f}"
        volume_por_hect = f"{row['Vol/Hect']:.3f}"
        diminuir_val = f"{row['diminuir']:.3f}"
        aumentar_val = f"{row['aumentar']:.3f}"
        dif_pct_val = f"{row['DAP %']:.2f}"  # Exibe a diferença percentual
        CAP_val = f"{row['mCAP']:.3f}"
        ALT_val = f"{row['mH']:.0f}"

        # Alterna as cores das linhas (índices pares e ímpares)
        tag = 'verde_claro' if i % 2 == 0 else 'branca'
        
        # Insere os dados na Treeview com a tag de cor
        table_ut_vol.insert("", "end", values=(ut_val, hectares_val, n_Árvores,
                                                Vol, Vol_Max, diminuir_val,
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
        
        #adicionando a coluna DAP_a
        df_saida = adicionarDAP_a()

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

            # Se o nome não estiver no dicionário de parâmetros, retorna a categoria original
            if nome not in parametros:
                return row["Categoria"]

            especie_parametros = parametros[nome]
            DAPmin = especie_parametros["dap_min"]  # Valor original de DAPmin
            DAPmax = especie_parametros["dap_max"]
            QF = especie_parametros["qf"]
            alt = especie_parametros["alt"]

            # Inicializa o DAPmin para a espécie em UT, caso haja edição
            DAPminEspUt = DAPmin
            QFEspUt = QF 
            # Verifica se há edições salvas para a UT e espécie no dicionário de dados editados
            if ut in dados_editados_por_ut and nome in dados_editados_por_ut[ut]:
                edits = dados_editados_por_ut[ut][nome]

                # Se "REM" for "SIM", marca como REMANESCENTE diretamente
                if edits.get("REM", "").strip().upper() == "SIM":
                    return "REMANESCENTE"

                # Substitui DAPmin com a edição de "DAP <" se estiver presente
                if "CAP <" in edits:
                    try:
                        DAPminEspUt = (float(edits["CAP <"]) / math.pi ) / 100  # DAPmin agora será o valor de edição específico para a UT
                        print("valor de  DAPminEspU")
                        print(DAPminEspUt)
                    except ValueError:
                        pass  # Se não for um número válido, mantém o valor original

                # Se houver edições para QF, substitui o valor de QF
                if "QF >=" in edits:
                    try:
                        QFEspUt = int(edits["QF >="])
                    except ValueError:
                        pass  # Se não for um número válido, mantém o valor original

            # Protegidas continuam sendo REMANESCENTE
            situacao = str(row.get("Situacao", "")).strip().lower()
            if situacao == "protegida":
                return "REMANESCENTE"

            # Regras de classificação de DAP
            if isinstance(row["DAP"], (float, int)):
                # Verifica se a coluna 'DAP_a' existe, e usa o valor de DAP_a se disponível
                dap = row["DAP_a"] if "DAP_a" in row and isinstance(row["DAP_a"], (float, int)) else row["DAP"]
                print("dap_a")
                print(dap)
                
                if dap < DAPmin:
                    return "REMANESCENTE"  # Se DAP estiver abaixo de DAPminEspUt, é REMANESCENTE
                elif DAPmin <= dap <= DAPminEspUt:
                    return "SUBSTITUTA"  # Se DAP estiver dentro do intervalo de DAPminEspUt e DAPmax, é SUBSTITUTA
                elif dap >= DAPmax:
                    return "REMANESCENTE"  # Se DAP for maior ou igual a DAPmax, também classifica como REMANESCENTE
                
            # Regras de classificação de QF
            if isinstance(row["QF"], int):
                if row["QF"] >= QF:
                    return "REMANESCENTE"
                elif QF > row["QF"] >= QFEspUt:
                    return "SUBSTITUTA"  # Se QF for menor que o valor de QF, é SUBSTITUTA


            
            # Se a altura for maior que o valor de 'alt', marca como REMANESCENTE
            if isinstance(row["H"], (float, int)) and alt > 0 and row["H"] > alt:
                return "REMANESCENTE"

            # Se nenhuma condição for atendida, retorna a categoria original
            return row["Categoria"]

        # Recuperar os parâmetros da tabela
        parametros_tabela = extrair_parametros_tabela()  # Função que recupera os parâmetros definidos

        # Aplicando a função ao DataFrame
        df_saida["Categoria"] = df_saida.apply(lambda row: filtrar_REM(row, parametros_tabela), axis=1)


        print()

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

    
        # **Contar a quantidade de árvores "CORTE" por UT e Nome Vulgar**
        df_contagem_corte = df_saida[df_saida["Categoria"] == "CORTE"].groupby(["UT", "Nome Vulgar"]).size().reset_index(name="Qtd_Corte")

        # **Mesclar com a contagem de cortes para verificar onde não há mais cortes**
        df_verificacao = df_saida[df_saida["Categoria"] == "SUBSTITUTA"].merge(df_contagem_corte, on=["UT", "Nome Vulgar"], how="left").fillna(0)

        # **Criar um identificador para marcar onde NÃO EXISTE "CORTE" (quando Qtd_Corte == 0)**
        df_verificacao["Marcar_REM"] = df_verificacao["Qtd_Corte"] == 0

        # **Criar uma lista de tuplas (UT, Nome Vulgar) onde as substitutas precisam virar REMANESCENTE**
        remover_tuplas = df_verificacao[df_verificacao["Marcar_REM"]][["UT", "Nome Vulgar"]].apply(tuple, axis=1).tolist()

        # **Atualizar df_saida para transformar "SUBSTITUTA" em "REMANESCENTE" onde não há cortes**
        df_saida["Categoria"] = df_saida.apply(
            lambda row: "REMANESCENTE" if (row["UT"], row["Nome Vulgar"]) in remover_tuplas and row["Categoria"] == "SUBSTITUTA" else row["Categoria"],
            axis=1
        )

        
        # **Verificando a contagem de cada categoria após a atualização**
        print("Contagem por Categoria após atualização:")
        print(df_saida["Categoria"].value_counts())

        # Exibindo a quantidade de "REMANESCENTE" e "SUBSTITUTA" para conferirmos o efeito
        print(df_saida[df_saida["Categoria"] == "REMANESCENTE"].shape[0], "REMANESCENTES")
        print(df_saida[df_saida["Categoria"] == "SUBSTITUTA"].shape[0], "SUBSTITUTAS")

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
        
        finalProcesso = time.time()
        print(f"Processamento realizado em {finalProcesso - inicioProcesso:.2f} s")
        
        
        #organizando as colunas
        df_saida = df_saida[COLUNAS_SAIDA]

        if contagem_categorias.get("CORTE", 0) > 0:
            print(f"CORTE: {contagem_categorias['CORTE']}")
            print(f"SUBSTITUTA: {contagem_categorias.get('SUBSTITUTA', 0)}")
            print(f"REMANESCENTE: {contagem_categorias.get('REMANESCENTE', 0)}")
            
            df_saida = ajustarVolumeHect()
            tabelaDeResumo()

            inicioTimeSalvar = time.time()
            print("dentro de processo")
            print(df_saida)
            if save == True :
                # Exporta para Excel usando o df_modificado com os cálculos originais
                # Exemplo de DataFrame a ser exportado
                df_export_planilha_principal = df_saida[["UT", "Faixa", "Placa", "Nome Vulgar", "CAP", "CAP_a", "mCAP", "H", "H_a", "mH", "QF",
                                    "X", "Y", "DAP", "DAP_a", "Vol", "Vol_a", "Lat", "Long", "DM", "OBS", "Categoria", "Situacao"]]
                
                df_export_tabela_de_ajuste = df_tabelaDeAjusteVol[['ut', 'Hectares', 'n_Árvores', 'Vol', 'Vol_Max', 'diminuir', 'aumentar', 'Vol/Hect', 'DAP %',  'mCAP', 'mH']]

                df_export_resumo = df_resumo[[
                                    "CAPmin", "Espécie", "n° Árvores", "Vol", "Vol/Árvore", 
                                    "Volume_a", "Vol_a/Árvore", "Vol/Hect"
                                ]]

                # Diretório onde o arquivo será salvo (exemplo)
                diretorio = os.path.dirname(entrada1_var.get())
                arquivo_saida = os.path.join(diretorio, "Planilha Processada - Handroanthus 1.0.xlsx")

                # Usando ExcelWriter para salvar o DataFrame em múltiplas abas
                with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
                    # Escreve o DataFrame na aba "Inventário Processado"
                    df_export_planilha_principal.to_excel(writer, sheet_name='Inventário Processado', index=False)
                    
                    # Escreve o DataFrame novamente na aba "ATabela de Ediçãoa"
                    df_export_tabela_de_ajuste.to_excel(writer, sheet_name='Resumo UT', index=False)

                    df_export_resumo.to_excel(writer, sheet_name="Resumo Espécie",  index=False)
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
            n_Árvores=('Nome Vulgar', 'count'),
            Vol=('Vol_a', 'sum'),
            CAPmin=('CAP_a', 'min')  # CAP mínimo
        ).reset_index()
        agrupado["vol_por_ha"] = agrupado["Vol"] / hectares

        

        nova_janela = tk.Toplevel()
        nova_janela.title(f"Espécies da UT {ut}")
        nova_janela.geometry("800x700")

        # Caminho para o ícone da janela
        icone_path = resource_path("src/img/icoGreenFlorest.ico")

        # Define o ícone da aplicação
        nova_janela.iconbitmap(icone_path)


        tabela = ttk.Treeview(nova_janela, columns=("CAPmin", "Nome", "n° Árvores", "Volume Total", "Vol/ha", "CAP <", "QF >=","CAP","H", "REM"),
                      show="headings", height=20, style="verde.Treeview")

        colunas = ("CAPmin", "Nome", "n° Árvores", "Volume Total", "Vol/ha", "CAP <", "QF >=","CAP","H", "REM")

        for col in colunas:
            tabela.heading(col, text=col)
            tabela.column(col, width=80, anchor="center")

        tabela.pack(pady=10, padx=10, fill="x")

        # Preenche a tabela com cores alternadas
        for i, row in agrupado.iterrows():
            nome = row["Nome Vulgar"]
            dados_salvos = dados_editados_por_ut.get(ut, {}).get(nome, {})

            valores = [
                f"{row['CAPmin']}",
                nome,
                row["n_Árvores"],
                f"{row['Vol']:.2f}",
                f"{row['vol_por_ha']:.2f}",
                dados_salvos.get("CAP <", ""),
                dados_salvos.get("QF >=", ""),
                dados_salvos.get("CAP", ""),
                dados_salvos.get("H", ""),
                dados_salvos.get("REM", "NÃO")
            ]

            tag = 'verde' if i % 2 == 0 else 'branco'
            tabela.insert("", "end", iid=nome, values=valores, tags=(tag,))

        tabela.tag_configure('verde', background="#e5fbe0")
        tabela.tag_configure('branco', background="#ffffff")
        # Adiciona o evento de clique duplo (Double Click) na coluna "Nome" (coluna 1)
        # Adiciona o evento de clique duplo (Double Click) na coluna "Nome" (coluna 2)]
            # Função que será chamada ao clicar no nome da espécie
        def ao_clicar_nome(event, ut):
           # Pega o item selecionado (linha)
            item = tabela.selection()[0]
            
            # Pega o valor da coluna "Nome" (coluna 1)
            nome = tabela.item(item, "values")[1]  # A coluna "Nome" está na posição 1

            # Aqui você pode chamar qualquer função com o nome e ut selecionados
            print(f"Clicou no nome: {nome}, UT: {ut}")

            # Filtrar df_saida pela espécie e UT com a categoria "CORTE"
            df_filtrado = df_saida[
                (df_saida["Nome Vulgar"] == nome) & 
                (df_saida["UT"] == int(ut)) & 
                (df_saida["Categoria"].str.upper() == "CORTE")
            ]

            # Verifique se o filtro está funcionando
            print(f"Linhas filtradas para {nome} na UT {ut}:")
            print(df_filtrado[['Nome Vulgar', 'CAP_a', 'H', 'QF', 'Vol', 'Categoria']])  # Exibe as colunas relevantes

            if df_filtrado.empty:
                print("Nenhum dado encontrado para esta espécie na UT selecionada.")
            
            # Ordenar os dados pela coluna "CAP" em ordem crescente
            df_filtrado = df_filtrado.sort_values(by="CAP_a", ascending=True)

            # Criar uma nova janela para exibir a tabela
            nova_janela = tk.Toplevel()
            nova_janela.title(f"Espécie: {nome} - UT {ut}")
            nova_janela.geometry("800x600")

            # Caminho para o ícone da janela
            icone_path = resource_path("src/img/icoGreenFlorest.ico")

            # Define o ícone da aplicação
            nova_janela.iconbitmap(icone_path)

            # Tabela com as colunas Nome Vulgar, CAP, H, QF, Vol, Categoria
            tabela_filtrada = ttk.Treeview(nova_janela, columns=("Nome Vulgar", "CAP", "H", "QF", "Vol", "Categoria"),
                                            show="headings", height=30)

            colunas = ["Nome Vulgar", "CAP", "H", "QF", "Vol", "Categoria"]

            for col in colunas:
                tabela_filtrada.heading(col, text=col)
                tabela_filtrada.column(col, width=100, anchor="center")

            # Adicionando a barra de rolagem
            scrollbar = ttk.Scrollbar(nova_janela, orient="vertical", command=tabela_filtrada.yview)
            tabela_filtrada.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")
            tabela_filtrada.pack(pady=10, padx=10, fill="x")

            # Alternar entre verde e branco
            tag = 'verde'
            
            # Preenche a tabela com as espécies filtradas
            for i, row in df_filtrado.iterrows():
                valores = [
                    row["Nome Vulgar"],
                    f"{row['CAP_a']}",
                    f"{row['H']}",
                    f"{row['QF']}",
                    f"{row['Vol']:.3f}",
                    row["Categoria"]
                ]
                
                # Gerar um iid único para cada linha com base no índice i
                iid_unico = f"{row['Nome Vulgar']}_{int(ut)}_{i}"  # Combine Nome Vulgar, UT e o índice para garantir unicidade
                
                # Alternância de cores, começa com "verde"
                if tag == 'verde':
                    tabela_filtrada.insert("", "end", iid=iid_unico, values=valores, tags=("verde",))
                    tag = 'branco'  # Alterna para "branco"
                else:
                    tabela_filtrada.insert("", "end", iid=iid_unico, values=valores, tags=("branco",))
                    tag = 'verde'  # Alterna para "verde"

            # Define as cores das tags para alternância de linhas
            tabela_filtrada.tag_configure('verde', background="#e5fbe0")
            tabela_filtrada.tag_configure('branco', background="#ffffff")


        tabela.bind("<Double-1>", lambda event: ao_clicar_nome(event, ut))
        # Área de edição
        frame_edicao = ttk.LabelFrame(nova_janela, text="Editar Espécie Selecionada")
        frame_edicao.pack(fill="x", padx=10, pady=10)

        # Criar 3 frames dentro de frame_edicao
        frame_substituta = ttk.Frame(frame_edicao)
        frame_substituta.grid(row=0, column=0, padx=10, pady=10)

        frame_ajustes = ttk.Frame(frame_edicao)
        frame_ajustes.grid(row=0, column=1, padx=10, pady=10)

        frame_remanescente = ttk.Frame(frame_edicao)
        frame_remanescente.grid(row=0, column=2, padx=10, pady=10)

        # Nomeando os frames
        ttk.Label(frame_substituta, text="Substituta :").grid(row=0, column=0, pady=5, sticky="w")
        ttk.Label(frame_ajustes, text="Ajuste :").grid(row=0, column=0, pady=5, sticky="w")
        ttk.Label(frame_remanescente, text="Remanescente :").grid(row=0, column=0, pady=5, sticky="w")

        # Dicionário para armazenar os widgets
        entradas = {}

        # Preencher o primeiro frame (Editar Substituta)
        for i, campo in enumerate(["CAP <", "QF >="]):  # A primeira lista de campos
            ttk.Label(frame_substituta, text=campo).grid(row=i+1, column=0, padx=5, pady=5, sticky="w")
            entry = ttk.Entry(frame_substituta, width=10)
            entry.grid(row=i+1, column=1, padx=5, pady=5)
            entradas[campo] = entry

        # Preencher o segundo frame (Editar Ajustes)
        for i, campo in enumerate(["CAP", "H"]):  # Outra lista de campos
            ttk.Label(frame_ajustes, text=campo).grid(row=i+1, column=0, padx=5, pady=5, sticky="w")
            entry = ttk.Entry(frame_ajustes, width=10)
            entry.grid(row=i+1, column=1, padx=5, pady=5)
            entradas[campo] = entry

        # Preencher o terceiro frame (REMANESCENTE)
        for i, campo in enumerate(["REM"]):  # Lista com apenas "REM"
            ttk.Label(frame_remanescente, text=campo).grid(row=i+1, column=0, padx=5, pady=5, sticky="w")
            combo = ttk.Combobox(frame_remanescente, values=["SIM", "NÃO"], state="readonly", width=10)
            combo.grid(row=i+1, column=1, padx=5, pady=5)
            entradas[campo] = combo

        # Alinhar todos os frames (substituta, ajustes, remanescente) de forma que fiquem com o mesmo tamanho e alinhados.
        frame_edicao.grid_columnconfigure(0, weight=1)
        frame_edicao.grid_columnconfigure(1, weight=1)
        frame_edicao.grid_columnconfigure(2, weight=1)

        # Ajustar alinhamento horizontal de "REM" com os outros campos
        frame_remanescente.grid_rowconfigure(0, weight=1)
        frame_remanescente.grid_rowconfigure(1, weight=1)


        
        especie_selecionada = tk.StringVar()

        def ao_selecionar_linha(event):
            item = tabela.selection()
            if item:
                nome = item[0]
                especie_selecionada.set(nome)
                dados = tabela.item(nome, "values")
                for i, campo in enumerate(colunas[5:]):
                    entradas[campo].delete(0, tk.END)
                    entradas[campo].insert(0, dados[i + 5])

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
            valores_atualizados = tabela.item(nome, "values")[:5] + tuple(entradas[campo].get() for campo in entradas)
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
        botao_trocar_tabela.config(text="Exibir Edição")
        tabela_visivel = False
    else:
        frame_resumo_especie.pack_forget()
        frame_tabela2.pack(fill="both", expand=True)
        botao_trocar_tabela.config(text="Exibir Resumo")
        tabela_visivel = True


# Função para garantir o caminho correto tanto no executável quanto no script de desenvolvimento
def resource_path(relative_path):
    """ Garante o caminho certo tanto no executável quanto em desenvolvimento """
    try:
        base_path = sys._MEIPASS  # Caso o script esteja executando como .exe (PyInstaller)
    except Exception:
        base_path = os.path.abspath(".")  # Caso esteja executando como script

    return os.path.join(base_path, relative_path)

# Criação da janela principal
app = tk.Tk()
app.title("Handroanthus 1.0")

# Caminho para o ícone da janela
icone_path = resource_path("src/img/icoGreenFlorest.ico")

# Define o ícone da aplicação
app.iconbitmap(icone_path)

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

style.configure("linha_verde.Treeview", background="#d3f8e2")  # verde clarinho
style.configure("linha_branca.Treeview", background="#ffffff")  # branco puro

# Calcular coordenadas para centralizar a janela
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

# Definir a geometria da janela com posição centralizada
app.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

# Impedir a maximização da janela
app.resizable(False, False)  # Permite redimensionamento
app.maxsize(largura_janela, altura_janela)  # Impede que a janela seja maximizada

# Caminho da imagem de fundo
caminho_imagem = resource_path("src/img/01florest.png")

# Verificar se a imagem existe
if not os.path.exists(caminho_imagem):
    print(f"Erro: Arquivo {caminho_imagem} não encontrado!")

# Abre a imagem e ajusta seu tamanho
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
    text="Modificar Remanescente",
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
    text="Exibir Resumo",  # ou "Mostrar Tabela de Ajustes" dependendo do estado atual
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

colunas_selecionados = ("Nome", "DAP <", "DAP >=", "QF >= ", "H >")

# Cria a Treeview
table_selecionados = ttk.Treeview(
    frame_listbox, columns=colunas_selecionados, show="headings", height=15
)

# Configura os títulos e colunas
for col in colunas_selecionados:
    table_selecionados.heading(col, text=col)
    table_selecionados.column(col, width=50, anchor="center")
table_selecionados.column("Nome", width=200, anchor="w")

# Scrollbar vertical
scrollbar_y = ttk.Scrollbar(frame_listbox, orient="vertical", command=table_selecionados.yview)
table_selecionados.configure(yscrollcommand=scrollbar_y.set)

# Posiciona na grid
table_selecionados.grid(row=1, column=1, padx=10, pady=10)
scrollbar_y.grid(row=1, column=2, sticky="ns")

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
    "CAP_min","Espécie", "n° Árvores","Vol", "Vol/Árvore", "Vol_a",
    "Vol/Árvore_a", "Vol/ha"
)

table_resumo_especie = ttk.Treeview(
    frame_resumo_especie,
    columns=colunas_resumo_especie,
    show="headings",
    height=22
)
for col in colunas_resumo_especie:
    table_resumo_especie.heading(col, text=col)
    table_resumo_especie.column(col, width=100, anchor="center")

# Scroll
scroll_resumo_especie = ttk.Scrollbar(frame_resumo_especie, orient="vertical", command=table_resumo_especie.yview)
table_resumo_especie.configure(yscrollcommand=scroll_resumo_especie.set)

# Packing
table_resumo_especie.pack(side="left", fill="both", expand=True)
scroll_resumo_especie.pack(side="right", fill="y")

# Associa as tags ao estilo
table_resumo_especie.tag_configure("linha_verde", background="#d3f8e2")
table_resumo_especie.tag_configure("linha_branca", background="#ffffff")

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
