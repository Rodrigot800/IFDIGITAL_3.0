import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time
from pacotes.edicaoValorFiltro import  abrir_janela_valores_padroes,valor1,valor2,valor3,valor4
import os
import configparser
import numpy as np
config = configparser.ConfigParser()
config.read('config.ini')


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
    "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria", "Situacao","UT_AREA_HA"
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
        #mesclagem da ut_area_ha com ut

        df_saida = pd.merge(
            df_saida,
            planilha_principal[["UT_ID", "UT_AREA_HA"]].drop_duplicates(),
            left_on="UT",
            right_on="UT_ID",
            how="left",
            suffixes=("", "_principal")
        )


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


        # **Exibir os primeiros resultados**
        #print(df_contagem)
        
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

        # Ordenar por UT, QF (maior para menor) e Volume_m3 (menor para maior)
        df_filtrado.sort_values(by=["UT", "QF", "Volume_m3"], ascending=[True, False, True], inplace=True)

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
        tk.messagebox.showinfo("SUCESSO",f" Processamento realizado em {finalProcesso - inicioProcesso:.2f} segundos e o  arquivo salvo em {finalTimeSalvar - inicioTimeSalvar:.2f} s ")

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


largura_janela = 800
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
