import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time
from pacotes.edicaoValorFiltro import  abrir_janela_valores_padroes,valor1,valor2,valor3,valor4
from pacotes.ordemSubstituta import OrdenadorFrame 
import os
import configparser
import numpy as np

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
ordering_mode = "QF > Vol_m3"

df_tabelaDeAjusteVol = None
global df_valores_atualizados

# Colunas de entrada e saída
COLUNAS_ENTRADA = [
    "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
    "X", "Y", "DAP", "Volumes (m³)", "Latitude", "Longitude", "DM", "Observações","DAP_C"
]

COLUNAS_SAIDA = [
    "UT", "Faixa", "Placa", "Nome Vulgar", "Nome Cientifico", "CAP", "ALT", "QF", "X", "Y",
    "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria", "Situacao","UT_AREA_HA"
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
        status_label.pack(pady=10)

        planilha_principal = pd.read_excel(arquivo1, engine="openpyxl")
        colunas_existentes = [col for col in planilha_principal.columns if col in ["Nome Vulgar"]]
        if not colunas_existentes:
            raise ValueError("A planilha principal não possui a coluna 'Nome Vulgar'.")
        nomes_vulgares = sorted(planilha_principal["Nome Vulgar"].dropna().unique()) 
        atualizar_listbox_nomes("")  # Inicializa a Listbox com todos os nomes

        frame_listbox_e_tabela.pack(pady=10)
        frame_secundario.pack(pady=10)
    except Exception as e:
        tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha principal: {e}")
    finally:
        status_label.pack_forget()

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

    # Adiciona à tabela com os valores atualizados do config.ini
    valores_atualizados = (nome, dap_max, dap_min, qf, alt, cap)
    table_selecionados.insert("", "end", values=valores_atualizados)

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
    for i in range(listbox_nomes_vulgares.size()):
        nome = listbox_nomes_vulgares.get(i)
        if not any(table_selecionados.item(child, "values")[0] == nome for child in table_selecionados.get_children()):
            table_selecionados.insert("", "end", values=(nome, dap, dap, qf, alt, cap))

# Função para remover o último item da tabela
def remover_ultimo_selecionado():
    filhos = table_selecionados.get_children()
    if filhos:
        table_selecionados.delete(filhos[-1])

# Função para limpar todos os itens da tabela
def limpar_lista_selecionados():
    table_selecionados.delete(*table_selecionados.get_children())

def ajustarVolumeHect(df_saida):
    global df_tabelaDeAjusteVol, df_valores_atualizados
    import numpy as np
    import pandas as pd
    import os

    # Cria uma cópia do DataFrame original para não modificá-lo
    df_modificado = df_saida.copy()

    # Cria as colunas CAP_m e ALT_m com valores padrão: CAP_m = 1, ALT_m = 0
    df_modificado["CAP_m"] = 1
    df_modificado["ALT_m"] = 0

    # Se houver um DataFrame global com valores atualizados, incorpora os valores de CAP_m e ALT_m
    if 'df_valores_atualizados' in globals() and not df_valores_atualizados.empty:
        # Se as colunas CAP_m e ALT_m não existem em df_valores_atualizados, tenta copiá-las de CAP e ALT
        for col in ["CAP_m", "ALT_m"]:
            if col not in df_valores_atualizados.columns and col.replace("_m", "") in df_valores_atualizados.columns:
                df_valores_atualizados[col] = df_valores_atualizados[col.replace("_m", "")]
        # Realiza o merge usando a chave UT (do df_modificado) e "ut" (do df_valores_atualizados)
        df_modificado = df_modificado.merge(
            df_valores_atualizados[["ut", "CAP_m", "ALT_m"]],
            left_on="UT", right_on="ut", how="left", suffixes=("", "_novo")
        )
        # Para CAP_m e ALT_m, se houver valor atualizado (não NaN), utiliza-o; senão, mantém o padrão
        for col in ["CAP_m", "ALT_m"]:
            if f"{col}_novo" in df_modificado.columns:
                df_modificado[col] = df_modificado[f"{col}_novo"].combine_first(df_modificado[col])
        # Remove as colunas auxiliares geradas no merge
        cols_to_drop = [f"{col}_novo" for col in ["CAP_m", "ALT_m"]]
        df_modificado.drop(columns=cols_to_drop, inplace=True)
        # Remove a coluna "ut" proveniente do merge, se existir (já temos "UT")
        if "ut" in df_modificado.columns and "UT" in df_modificado.columns:
            df_modificado.drop(columns=["ut"], inplace=True)

    # Cria as colunas "ut" e "hac" (para facilitar os agrupamentos)
    df_modificado["ut"] = df_modificado["UT"]
    df_modificado["hac"] = df_modificado["UT_AREA_HA"]

    # Cria as colunas de cálculo: ALT_C e CAP_C e calcula DAP_C a partir de CAP_C
    df_modificado["ALT_C"] = df_modificado["ALT"]
    df_modificado["CAP_C"] = df_modificado["CAP"]
    df_modificado["DAP_C"] = (df_modificado["CAP_C"] / np.pi) / 100

    # Converte para numérico (garante que DAP_C e ALT sejam float)
    df_modificado["DAP_C"] = pd.to_numeric(df_modificado["DAP_C"], errors='coerce')
    df_modificado["ALT"] = pd.to_numeric(df_modificado["ALT"], errors='coerce')

    # Cálculo do Volume_m3_C
    df_modificado["Volume_m3_C"] = ((df_modificado["DAP_C"] ** 2) * np.pi / 4) * df_modificado["ALT"] * 0.7

    # Normaliza a coluna "Categoria"
    df_modificado["Categoria"] = df_modificado["Categoria"].astype(str).str.strip().str.lower()

    # Filtra os registros com Categoria "corte"
    df_filtrado = df_modificado[df_modificado["Categoria"].isin(["corte"])]

    if df_filtrado.empty:
        print("Nenhuma árvore foi categorizada como 'corte'.")
        # Mantemos CAP_m e ALT_m na seleção para persistência
        df_tabelaDeAjusteVol = df_modificado[["ut", "hac", "CAP_m", "ALT_m"]].drop_duplicates()
        df_tabelaDeAjusteVol["num_arvores"] = 0
        df_tabelaDeAjusteVol["volume_total"] = 0
        df_tabelaDeAjusteVol["diminuir"] = 0
        df_tabelaDeAjusteVol["aumentar"] = 0
    else:
        # Agrupa para contar o número de árvores e somar o volume total por UT
        contagem_arvores = df_filtrado.groupby("ut").size().reset_index(name="num_arvores")
        volume_total_por_ut = df_filtrado.groupby("ut")["Volume_m3_C"].sum().reset_index()
        volume_total_por_ut.rename(columns={"Volume_m3_C": "volume_total"}, inplace=True)
        
        # Cria o DataFrame de ajuste com UT, Hectares, CAP_m e ALT_m para preservar os valores de edição
        df_tabelaDeAjusteVol = df_modificado[["ut", "hac", "CAP_m", "ALT_m"]].drop_duplicates()
        print("Dados iniciais de UT, Hectares, CAP_m e ALT_m:")
        print(df_tabelaDeAjusteVol)

        # Calcula o volume máximo como (hectares * 30)
        df_tabelaDeAjusteVol["volume_max"] = df_tabelaDeAjusteVol["hac"] * 30

        # Faz merge para incorporar a contagem de árvores e o volume total
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(contagem_arvores, on="ut", how="left").fillna(0)
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(volume_total_por_ut, on="ut", how="left").fillna(0)

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

    # Converte o número de árvores para inteiro
    df_tabelaDeAjusteVol["num_arvores"] = df_tabelaDeAjusteVol["num_arvores"].astype(int)

    # Garante que a coluna DAP_t exista
    if "DAP_t" not in df_tabelaDeAjusteVol.columns:
        df_tabelaDeAjusteVol["DAP_t"] = 0

    # Se houver df_valores_atualizados, já foram mesclados anteriormente.
    # Caso contrário, os padrões em df_modificado já foram definidos (CAP_m=1, ALT_m=0)
    # Aqui, opcionalmente, você pode realizar outro merge para garantir que os valores persistam.
    if 'df_valores_atualizados' in globals() and not df_valores_atualizados.empty:
        # Se as colunas CAP_m e ALT_m não existem, já foram criadas no início (padrão 1 e 0)
        df_val = df_valores_atualizados[["ut", "CAP_m", "ALT_m"]].drop_duplicates(subset="ut")
        df_tabelaDeAjusteVol = df_tabelaDeAjusteVol.merge(df_val, on="ut", how="left", suffixes=("", "_novo"))
        # Se existir valor atualizado (não NaN) para CAP_m e ALT_m, usa-o; senão, mantém o já existente
        for col in ["CAP_m", "ALT_m"]:
            if f"{col}_novo" in df_tabelaDeAjusteVol.columns:
                df_tabelaDeAjusteVol[col] = df_tabelaDeAjusteVol[f"{col}_novo"].combine_first(df_tabelaDeAjusteVol[col])
        # Remove as colunas auxiliares
        cols_to_drop = [f"{col}_novo" for col in ["CAP_m", "ALT_m"]]
        df_tabelaDeAjusteVol.drop(columns=cols_to_drop, inplace=True)
        # Preenche NaN com os padrões desejados (para UTs sem atualização)
        df_tabelaDeAjusteVol["CAP_m"] = df_tabelaDeAjusteVol["CAP_m"].fillna(1)
        df_tabelaDeAjusteVol["ALT_m"] = df_tabelaDeAjusteVol["ALT_m"].fillna(0)
    else:
        df_tabelaDeAjusteVol["CAP_m"] = 1
        df_tabelaDeAjusteVol["ALT_m"] = 0

    # Limpa a Treeview table_ut_vol para evitar duplicação
    for child in table_ut_vol.get_children():
        table_ut_vol.delete(child)

    # Insere os valores na Treeview table_ut_vol
    for _, row in df_tabelaDeAjusteVol.iterrows():
        ut_val = f"{row['ut']:.0f}"
        hectares_val = f"{row['hac']:.5f}"
        num_arvores = f"{row['num_arvores']:.0f}"
        volume_total = f"{row['volume_total']:.5f}"
        volume_max = f"{row['volume_max']:.5f}"
        volume_por_hect = f"{row['volume_por_hectare']:.2f}"
        diminuir_val = f"{row['diminuir']:.3f}"
        aumentar_val = f"{row['aumentar']:.3f}"
        DAP_val = f"{row['DAP_t']:.2f}"
        CAP_val = f"{row['CAP_m']:.1f}"
        ALT_val = f"{row['ALT_m']:.1f}"
        table_ut_vol.insert("", "end", values=(ut_val, hectares_val, num_arvores,
                                                volume_total, volume_max, diminuir_val,
                                                aumentar_val, volume_por_hect,
                                                DAP_val, CAP_val, ALT_val))
    print("UT, Hectares, Número de Árvores e Volume Total Atualizados:")
    print(df_tabelaDeAjusteVol)

    # Exporta para Excel usando o df_modificado com os cálculos originais
    df_export = df_modificado[["UT", "Faixa", "Placa", "Nome Vulgar", "CAP", "CAP_C", "CAP_m", "ALT", "ALT_C", "ALT_m", "QF",
                               "X", "Y", "DAP", "DAP_C", "Volume_m3", "Volume_m3_C", "DM", "Categoria"]]
    diretorio = os.path.dirname(entrada1_var.get())
    arquivo_saida = os.path.join(diretorio, "Planilha ajustada - IFDIGITAL 3.0.xlsx")
    df_export.to_excel(arquivo_saida, index=False, engine="xlsxwriter")



def processar_planilhas():
    
     # Oculta o botão e exibe/inicia a barra de progresso
    button_process.pack_forget()
    progress_bar.pack(pady=10)
    progress_bar.start(10)
    app.update_idletasks()  # Garante que a interface seja atualizada


    inicioProcesso = time.time()
    """Processa os dados da planilha principal e mescla com nomes científicos."""
    global planilha_principal, especies_selecionados

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

       

        # Primeiro, marque como "REM" se a espécie não estiver selecionada.
        df_saida["Categoria"] = df_saida["Nome Vulgar"].apply(
            lambda nome: "REM" if nome.upper() not in nomes_selecionados else "CORTE"
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
                            "alt": float(valores[4]) if valores[4] else 0.0,
                            "cap": float(valores[5]) if valores[5] else 0.0
                        }
                    except ValueError as e:
                        print(f"Erro ao processar valores para {nome}: {e}")
                        continue
            return parametros

        # Função para aplicar o filtro
        def filtrar_REM(row, parametros):
            nome = row["Nome Vulgar"].upper()

            # Verificar se a espécie está na tabela de parâmetros
            if nome not in parametros:
                return row["Categoria"]  # Retorna a categoria atual se a espécie não estiver na tabela

            # Extrair os parâmetros da tabela para a espécie
            especie_parametros = parametros[nome]
            DAPmin = especie_parametros["dap_min"]
            DAPmax = especie_parametros["dap_max"]
            QF = especie_parametros["qf"]
            alt = especie_parametros["alt"]

            # Atualizar a coluna "Categoria" com "REM" se a situação for "protegida"
            situacao = str(row["Situacao"]).strip().lower() if pd.notna(row["Situacao"]) else ""

            # Se for protegida, marcar como REM
            if situacao == "protegida":
                return "REM"

            # Se o DAP for menor que DAPmin ou maior/igual a DAPmax, marcar como REM
            if isinstance(row["DAP"], (float, int)) and (row["DAP"] < DAPmin or row["DAP"] >= DAPmax):
                return "REM"

            # Se QF for igual ao valor definido, marcar como REM
            if isinstance(row["QF"], int) and row["QF"] == QF:
                return "REM"

            # Se alt > 0 e ALT for maior que alt, marcar como REM
            if isinstance(row["ALT"], (float, int)) and alt > 0 and row["ALT"] > alt:
                return "REM"

            # Se nenhuma condição foi atendida, mantém a categoria original
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
        def definir_sbustituta_vuneravel(quantidade, Situacao, area_hect):
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
            lambda row: definir_sbustituta_vuneravel(row["Quantidade"], row["Situacao"], row["UT_AREA_HA"]), 
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
        if ordering_mode == "QF > Vol_m3":
            df_filtrado.sort_values(by=["UT", "QF", "Volume_m3"], ascending=[True, False, True], inplace=True)
            print("-----QF > Vol_m3-----")
        elif ordering_mode == "Vol_m3 > QF":
            df_filtrado.sort_values(by=["UT", "Volume_m3", "QF"], ascending=[True, True, False], inplace=True)
            print("----------")
        elif ordering_mode == "Apenas QF":
            df_filtrado.sort_values(by=["UT", "QF"], ascending=[True, False], inplace=True)
            print("----------")
        elif ordering_mode == "Apenas Vol_m3":
            df_filtrado.sort_values(by=["UT", "Volume_m3"], ascending=[True, True], inplace=True)
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

        # **Criar uma lista de tuplas (UT, Nome Vulgar) onde as substitutas precisam virar REM**
        remover_tuplas = df_verificacao[df_verificacao["Marcar_REM"]][["UT", "Nome Vulgar"]].apply(tuple, axis=1).tolist()

        # **Atualizar df_saida para transformar "SUBSTITUTA" em "REM" onde não há cortes**
        df_saida["Categoria"] = df_saida.apply(
            lambda row: "REM" if (row["UT"], row["Nome Vulgar"]) in remover_tuplas and row["Categoria"] == "SUBSTITUTA" else row["Categoria"],
            axis=1
        )

        # **Verificar os resultados corrigidos**
        print("\n--- Linhas que viraram REM porque não há mais CORTE dentro da UT ---")
        print(df_saida[df_saida["Categoria"] == "REM"][["UT", "Nome Vulgar", "Categoria"]].drop_duplicates())

        print("\n--- Contagem Final por Categoria ---")
        print(df_saida["Categoria"].value_counts())


        contagem_categorias = df_saida["Categoria"].value_counts()

        print("Contagem por Categoria:")
        print(f"CORTE: {contagem_categorias.get('CORTE', 0)}")
        print(f"SUBSTITUTA: {contagem_categorias.get('SUBSTITUTA', 0)}")
        print(f"REM: {contagem_categorias.get('REM', 0)}")

        print(f"Numero total de linhas em df_saida: {len(df_saida)}")
        
        
        

        finalProcesso = time.time()
        print(f"Processamento realizado em {finalProcesso - inicioProcesso:.2f} s")

        inicioTimeSalvar = time.time()
        ajustarVolumeHect(df_saida)
        #organizando as colunas
        df_saida = df_saida[COLUNAS_SAIDA]
        #salvar o arquivo no diretório
        # diretorio = os.path.dirname(entrada1_var.get())
        # arquivo_saida = os.path.join(diretorio, "Planilha Processada - IFDIGITAL 3.0.xlsx")
        # df_saida.to_excel(arquivo_saida,index=False,engine="xlsxwriter")
        finalTimeSalvar = time.time()

        # finalizando barra de progresso
        # Para e esconde a barra de progresso e exibe novamente o botão
        progress_bar.stop()
        progress_bar.pack_forget()
        button_process.pack(pady=10)
        
        tk.messagebox.showinfo("SUCESSO",f" Processamento realizado em {finalProcesso - inicioProcesso:.2f} segundos e o  arquivo salvo em {finalTimeSalvar - inicioTimeSalvar:.2f} s ")

    except Exception as e:
        # finalizando barra de progresso
        # Para e esconde a barra de progresso e exibe novamente o botão
        progress_bar.stop()
        progress_bar.pack_forget()
        button_process.pack(pady=10)
        tk.messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")

def iniciar_processamento():
    config = configparser.ConfigParser()
    config.read('config.ini')
    """Inicia o processamento em uma thread separada."""
    thread = threading.Thread(
    target=processar_planilhas
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
colunas_editaveis = [ "CAP", "ALT"]

def editar_celula_volume(event):
    """
    Permite editar as colunas CAP e ALT ao dar duplo clique em uma célula da Treeview (table_ut_vol).
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
    # Verifica se a coluna é editável (lista global definida, por exemplo: colunas_editaveis = ["CAP", "ALT"])
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

        import pandas as pd
        import numpy as np

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

        # Se a coluna editada for CAP ou ALT, atualiza também o DataFrame global de valores atualizados
        if col_nome in ["CAP", "ALT"]:
            # Se df_valores_atualizados não existir ou estiver vazio, inicializa-o
            if 'df_valores_atualizados' not in globals() or df_valores_atualizados.empty:
                df_valores_atualizados = pd.DataFrame(columns=["ut", "CAP", "ALT"])
            # Procura se já existe uma linha para essa UT
            idx_val_series = df_valores_atualizados[df_valores_atualizados["ut"] == ut_val].index
            if not idx_val_series.empty:
                df_valores_atualizados.at[idx_val_series[0], col_nome] = novo_valor
            else:
                nova_linha = {"ut": ut_val, "CAP": np.nan, "ALT": np.nan}
                nova_linha[col_nome] = novo_valor
                df_valores_atualizados = pd.concat([df_valores_atualizados, pd.DataFrame([nova_linha])], ignore_index=True)

        print(f"Valor atualizado: UT={ut_val}, {col_nome}={novo_valor}")
        print("DataFrame de ajuste atualizado:")
        print(df_tabelaDeAjusteVol)
        print("DataFrame de valores atualizados:")
        print(df_valores_atualizados)

        entry.destroy()

    # Vincula a função salvar_novo_valor aos eventos Return e FocusOut
    entry.bind("<Return>", salvar_novo_valor)
    entry.bind("<FocusOut>", salvar_novo_valor)



# Interface gráfica
app = tk.Tk()

app.title("IFDIGITAL 3.0")
app.geometry("1200x1200")


largura_janela = 1500
altura_janela = 900

# Obter largura e altura da tela
largura_tela = app.winfo_screenwidth()
altura_tela = app.winfo_screenheight()

# Calcular coordenadas para centralizar a janela
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

# Definir a geometria da janela com posição centralizada
app.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

entrada1_var = tk.StringVar()
entrada2_var = tk.StringVar()
pesquisa_var = tk.StringVar()

# Frame para entrada de arquivos
frame_inputs = ttk.LabelFrame(app, text="Entrada de Arquivos", padding=(10, 10))
frame_inputs.pack(fill="x", pady=10, padx=10)

ttk.Label(frame_inputs, text="INVENTÁRIO :").grid(row=0, column=0, sticky="w")
ttk.Entry(frame_inputs, textvariable=entrada1_var, width=60).grid(row=0, column=1, pady=5, padx=5)
ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("principal")).grid(row=0, column=2, padx=5)

ttk.Label(frame_inputs, text="LISTA DE ESPECIES :").grid(row=1, column=0, sticky="w")
ttk.Entry(frame_inputs, textvariable=entrada2_var, width=60).grid(row=1, column=1, pady=5, padx=5)
ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("secundaria")).grid(row=1, column=2, padx=5)
#ttk.Button(frame_inputs, text="Novo Botão", command=lambda:abrir_janela_valores_padroes(root) ).grid(row=2, column=1, pady=10, padx=5)  # Botão abaixo dos inputs

frame_lado_a_lado = tk.Frame(app)
frame_lado_a_lado.pack(pady=10)

botao_modificar_filtro = tk.Button(frame_lado_a_lado, text="Modificar Filtragem para Substituta", 
                                    command=abrir_janela_valores_padroes_callback)
botao_modificar_filtro.pack(side=tk.LEFT, padx=5)

combobox = ttk.Combobox(frame_lado_a_lado, state="readonly", width=30,
                        values=[
                            "QF > Vol_m3",
                            "Vol_m3 > QF",
                            "Apenas QF",
                            "Apenas Vol_m3"
                        ])
combobox.current(0)  # Seleciona a primeira opção por padrão
combobox.bind("<<ComboboxSelected>>", update_ordering_mode)
combobox.pack(side=tk.LEFT, padx=5)

# Criação dos widgets que serão atualizados
status_label = ttk.Label(app, text="")  # Inicialmente vazio
status_label.pack(pady=10)

frame_listbox_e_tabela =tk.Frame(app)
frame_listbox_e_tabela.pack(side="left", padx=10, pady=10, fill="y")

# Frame para Listboxes

frame_listbox = ttk.LabelFrame(frame_listbox_e_tabela, text="Seleção de Nomes Vulgares", padding=(10, 10))
frame_listbox.pack(side="left", padx=10, pady=10, fill="y")


# Criando um Frame para alinhar a Label e a Entry horizontalmente
frame_pesquisa = tk.Frame(frame_listbox)
frame_pesquisa.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")

# Adicionando a Label e a Entry dentro do frame_pesquisa
ttk.Label(frame_pesquisa, text="Pesquisar:").pack(side="left")
pesquisa_entry = ttk.Entry(frame_pesquisa, textvariable=pesquisa_var, width=40)
pesquisa_entry.pack(side="left", padx=5)

# Vinculando evento de pesquisa
pesquisa_entry.bind("<KeyRelease>", pesquisar_nomes)

colunas_selecionados = ("Nome", "DAP <", "DAP >=", "QF = ", "ALT >", "CAP <")

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

# Botões
btn_todos = ttk.Button(frame_listbox, text="Selecionar Todos", command=selecionar_todos)
btn_todos.grid(row=2, column=0, padx=5, pady=10)

btn_remover = ttk.Button(frame_listbox, text="Remover Último", command=remover_ultimo_selecionado)
btn_remover.grid(row=2, column=1, padx=5, pady=10)

btn_limpar = ttk.Button(frame_listbox, text="Limpar Lista", command=limpar_lista_selecionados)
btn_limpar.grid(row=2, column=2, padx=5, pady=10)

#frma para a tabela de ajuste de volume por hectar 

# Segunda tabela com "UT" e "Vol/Hect Total"
frame_tabela2 = ttk.Frame(frame_listbox_e_tabela)
frame_tabela2.pack(side="right", padx=10, pady=10, fill="y")

# Definição das colunas
colunas_tabela2 = ("UT", "Hectares", "n° Árv", "Vol(m³)", "Vol_Max", "Diminuir", "Aumentar", "V_m³/ha", "DAP", "CAP", "ALT")
table_ut_vol = ttk.Treeview(frame_tabela2, columns=colunas_tabela2, show="headings", height=5)

# Configuração das colunas
for col in colunas_tabela2:
    table_ut_vol.heading(col, text=col)
    table_ut_vol.column(col, width=70, anchor="center")

table_ut_vol.pack(pady=10)
table_ut_vol.config(height=20)

# Adiciona evento de duplo clique para editar
table_ut_vol.bind("<Double-1>", editar_celula_volume)

# Botão abaixo da segunda tabela
btn_confirmar = ttk.Button(frame_tabela2, text="editar cap alt", command=lambda: print("Dados confirmados!"))
btn_confirmar.pack(pady=10)

# Frame para o botão de processamento
frame_secundario = ttk.Frame(app, padding=(10, 10))
button_process = ttk.Button(frame_secundario, text="Processar Planilhas", command=iniciar_processamento, width=40)
button_process.pack(pady=10)

# Barra de progresso (inicialmente não exibida)
progress_bar = ttk.Progressbar(frame_secundario, mode='indeterminate', length=300)
frame_listbox_e_tabela.pack_forget()
frame_secundario.pack_forget()  # Inicialmente escondido


carregar_planilha_salva("principal")
carregar_planilha_salva("secundaria")

app.mainloop()
