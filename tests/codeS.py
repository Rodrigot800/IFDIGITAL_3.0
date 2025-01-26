
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pacotes.view.interface import InterfaceApp

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

def carregar_planilha_principal(self, arquivo):
    """Carrega a planilha principal e exibe os nomes vulgares."""
    try:
        # Atualiza o rótulo para exibir o status de carregamento
        self.status_label.config(text="Carregando inventário principal...")
        self.status_label.pack(pady=10)

        # Carrega a planilha principal
        planilha_principal = pd.read_excel(arquivo, engine="openpyxl")
        
        # Verifica se a coluna "Nome Vulgar" está presente
        if "Nome Vulgar" not in planilha_principal.columns:
            raise ValueError("A planilha principal não possui a coluna 'Nome Vulgar'.")

        # Extrai nomes vulgares únicos, ignorando valores nulos
        self.nomes_vulgares = sorted(planilha_principal["Nome Vulgar"].dropna().unique())
        
        # Atualiza a Listbox com os nomes vulgares
        self.atualizar_listbox_nomes("")
        
        # Exibe os frames de seleção e processamento
        self.frame_listbox.pack(pady=10)
        self.frame_secundario.pack(pady=10)

    except Exception as e:
        # Exibe mensagem de erro caso algo dê errado
        tk.messagebox.showerror("Erro", f"Falha ao carregar a planilha principal: {e}")
    finally:
        # Esconde o rótulo de status
        self.status_label.pack_forget()



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
        df2.rename(columns={
            "NOME_VULGAR": "Nome Vulgar", 
            "NOME_CIENTIFICO": "Nome Cientifico",
            "SITUACAO": "Situacao"
        }, inplace=True)

        # Normalizar as colunas "Nome Vulgar" e "Nome Cientifico" e mesclar os dados
        df_saida["Nome Vulgar"] = df_saida["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Vulgar"] = df2["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Cientifico"] = df2["Nome Cientifico"].str.strip().str.upper()
        df_saida = pd.merge(df_saida, df2[["Nome Vulgar", "Nome Cientifico", "Situacao"]], 
                            on="Nome Vulgar", how="left")
        #categorizando as arvores REM
        # Atualizar a coluna "Categoria" com "REM" se a situação for "protegida"
        df_saida.loc[df_saida["Situacao"].str.lower() == "protegida", "Categoria"] = "REM"
        # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'DAP' < 0.5
        df_saida.loc[df_saida["DAP"] < 0.5, "Categoria"] = "REM"
        # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'DAP' >= 2
        df_saida.loc[df_saida["DAP"] >= 2, "Categoria"] = "REM"
        # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'QF' = 3
        df_saida.loc[df_saida["QF"] == 3, "Categoria"] = "REM"
        # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'ALT' > (DEFINITO PELO USUARIO)
        # Criterio e opcional
        # df_saida.loc[df_saida["ALT"] >= 0, "Categoria" ] = "REM"



        # #apagar acoluna "situação após a ultilizacao"
        # df_saida = df_saida.drop(columns="Situacao")

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
if __name__ == "__main__":
    root = tk.Tk()
    app = InterfaceApp(root)
    root.mainloop()


#2.0
# import pandas as pd
# import tkinter as tk
# from tkinter import filedialog, messagebox, ttk
# import threading
# import os
# import time
# from pacotes.ajustar_largura_colunas import ajustar_largura_colunas

# # Colunas de entrada e saída
# COLUNAS_ENTRADA = [
#     "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
#     "X", "Y", "X Corrigido", "Y Corrigido", "DAP", "Volumes (m³)", "X Negativo", "Y Negativo",
#     "Latitude", "Longitude", "DM", "Observações", "N"
# ]

# COLUNAS_SAIDA = [
#     "UT", "Faixa", "Placa", "Nome Vulgar", "CAP", "ALT", "QF", "X", "Y",
#     "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria", "Nome Cientifico"
# ]

# # Variável para armazenar a planilha carregada
# planilha_principal = None

# def selecionar_arquivos():
#     """Seleciona os arquivos das duas planilhas."""
#     global planilha_principal

#     arquivo1 = filedialog.askopenfilename(
#         title="Selecione a planilha de entrada (dados principais)",
#         filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
#     )
#     if arquivo1:
#         entrada1_var.set(arquivo1)
#         # Carregar a planilha de forma assíncrona
#         threading.Thread(target=carregar_planilha_principal, args=(arquivo1,)).start()

#     arquivo2 = filedialog.askopenfilename(
#         title="Selecione a segunda planilha (Nomes Vulgares e Científicos)",
#         filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
#     )
#     if arquivo2:
#         entrada2_var.set(arquivo2)

# def carregar_planilha_principal(arquivo1):
#     """Carrega a planilha principal em segundo plano."""
#     global planilha_principal
#     try:
#         print("Carregando a planilha principal...")
#         planilha_principal = pd.read_excel(arquivo1, engine="openpyxl")
#         colunas_existentes = [col for col in planilha_principal.columns if col in COLUNAS_ENTRADA]
#         if not colunas_existentes:
#             raise ValueError("A planilha principal não possui as colunas esperadas.")
#         planilha_principal = planilha_principal[colunas_existentes]
#         print("Planilha principal carregada com sucesso.")
#     except Exception as e:
#         messagebox.showerror("Erro", f"Falha ao carregar a planilha principal: {e}")

# def processar_planilhas():
#     """Processa os dados da planilha principal e mescla com nomes científicos."""
#     global planilha_principal

#     arquivo2 = entrada2_var.get()

#     # Verificar se a planilha principal já foi carregada
#     if planilha_principal is None:
#         messagebox.showerror("Erro", "Por favor, selecione e aguarde o carregamento da planilha principal.")
#         return

#     # Verificar se o segundo arquivo foi selecionado
#     if not arquivo2:
#         messagebox.showerror("Erro", "Por favor, selecione o arquivo de Nomes Vulgares e Científicos.")
#         return

#     try:
#         start_time = time.time()
#         progress_var.set(10)  # Atualiza o progresso inicial
#         progress_bar.pack(pady=10, fill="x")  # Mostra a barra de progresso
#         progress_bar.update()

#         # Criar o DataFrame de saída
#         print("Criando o DataFrame de saída...")
#         df_saida = pd.DataFrame()
#         for entrada, saida in {
#             "UT": "UT",
#             "Faixa": "Faixa",
#             "Placa": "Placa",
#             "Nome Vulgar": "Nome Vulgar",
#             "CAP": "CAP",
#             "ALT": "ALT",
#             "QF": "QF",
#             "X": "X",
#             "Y": "Y",
#             "DAP": "DAP",
#             "Volumes (m³)": "Volume_m3",
#             "Latitude": "Latitude",
#             "Longitude": "Longitude",
#             "DM": "DM",
#             "Observações": "Observacoes"
#         }.items():
#             if entrada in planilha_principal.columns:
#                 df_saida[saida] = planilha_principal[entrada]
#             else:
#                 df_saida[saida] = None  # Colunas ausentes preenchidas com None
#         if "Categoria" not in df_saida.columns:
#             df_saida["Categoria"] = None
#         progress_var.set(40)
#         progress_bar.update()

#         # Carregar a segunda planilha
#         print("Carregando a segunda planilha...")
#         df2 = pd.read_excel(arquivo2, engine="openpyxl")
#         df2.rename(columns={"NOME_VULGAR": "Nome Vulgar", "NOME_CIENTIFICO": "Nome Cientifico"}, inplace=True)

#         # Normalizar as colunas "Nome Vulgar" e mesclar os dados
#         df_saida["Nome Vulgar"] = df_saida["Nome Vulgar"].str.strip().str.upper()
#         df2["Nome Vulgar"] = df2["Nome Vulgar"].str.strip().str.upper()
#         df_saida = pd.merge(df_saida, df2[["Nome Vulgar", "Nome Cientifico"]], on="Nome Vulgar", how="left")
#         progress_var.set(80)
#         progress_bar.update()

#         # Reordenar as colunas para garantir que "Nome Cientifico" esteja ao lado de "Nome Vulgar"
#         colunas = list(df_saida.columns)
#         if "Nome Cientifico" in colunas:
#             colunas.remove("Nome Cientifico")
#             idx = colunas.index("Nome Vulgar") + 1
#             colunas.insert(idx, "Nome Cientifico")
#         df_saida = df_saida[colunas]
        
#         # Salvar o arquivo de saída
#         print("Salvando o arquivo de saída...")
#         diretorio = os.path.dirname(entrada1_var.get())
#         arquivo_saida = os.path.join(diretorio, "planilha_processada_completa.xlsx")
#         df_saida.to_excel(arquivo_saida, index=False, engine="xlsxwriter")
#         progress_var.set(100)
#         progress_bar.update()

#         # Após salvar o arquivo
#         # ajustar_largura_colunas(arquivo_saida)

#         elapsed_time = time.time() - start_time
#         print(f"Processamento concluído em {elapsed_time:.2f} segundos.")
#         messagebox.showinfo("Sucesso", f"Planilha processada salva em:\n{arquivo_saida}\nTempo total: {elapsed_time:.2f} segundos")
#     except Exception as e:
#         messagebox.showerror("Erro", f"Falha ao processar as planilhas: {e}")
#     finally:
#         progress_var.set(0)  # Resetar a barra de progresso
#         progress_bar.pack_forget()  # Esconde a barra de progresso
#         progress_bar.update()

# def iniciar_processamento():
#     """Inicia o processamento em uma thread separada."""
#     thread = threading.Thread(target=processar_planilhas)
#     thread.daemon = True  # Fecha a thread quando a interface é fechada
#     thread.start()

# # Interface gráfica
# app = tk.Tk()
# app.title("Processador de Inventário e Mesclagem")
# app.geometry("600x400")

# entrada1_var = tk.StringVar()
# entrada2_var = tk.StringVar()
# progress_var = tk.IntVar()

# tk.Label(app, text="Arquivo 1: Planilha Principal").pack(pady=5)
# tk.Entry(app, textvariable=entrada1_var, width=60).pack(pady=5)
# tk.Label(app, text="Arquivo 2: Nomes Vulgares e Científicos").pack(pady=5)
# tk.Entry(app, textvariable=entrada2_var, width=60).pack(pady=5)

# progress_bar = ttk.Progressbar(app, variable=progress_var, maximum=100)

# tk.Button(app, text="Selecionar Planilhas", command=selecionar_arquivos).pack(pady=10)
# tk.Button(app, text="Processar e Mesclar Planilhas", command=iniciar_processamento, bg="green", fg="white").pack(pady=20)

# app.mainloop()




# import pandas as pd
# import tkinter as tk
# from tkinter import filedialog, messagebox, ttk
# import threading
# import os
# import time

# # Variáveis globais
# planilha_principal = None
# planilha_secundaria = None
# nomes_vulgares = []  # Lista de todos os nomes vulgares
# nomes_selecionados = []  # Lista para manter a ordem dos nomes selecionados

# # Colunas de entrada e saída
# COLUNAS_ENTRADA = [
#     "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
#     "X", "Y", "DAP", "Volumes (m³)", "Latitude", "Longitude", "DM", "Observações"
# ]

# COLUNAS_SAIDA = [
#     "UT", "Faixa", "Placa", "Nome Vulgar",  "Nome Cientifico" , "CAP", "ALT", "QF", "X", "Y",
#     "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria"
# ]

# # Funções da Interface e Processamento

# def selecionar_arquivos(tipo):
#     """Seleciona os arquivos das planilhas."""
#     global planilha_principal, planilha_secundaria

#     arquivo = filedialog.askopenfilename(
#         title=f"Selecione a planilha {'principal' if tipo == 'principal' else 'secundária'}",
#         filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
#     )
#     if arquivo:
#         if tipo == "principal":
#             entrada1_var.set(arquivo)
#             threading.Thread(target=carregar_planilha_principal, args=(arquivo,)).start()
#         elif tipo == "secundaria":
#             entrada2_var.set(arquivo)
#             threading.Thread(target=carregar_planilha_secundaria, args=(arquivo,)).start()


# def carregar_planilha_principal(arquivo1):
#     """Carrega a planilha principal e exibe os nomes vulgares."""
#     global planilha_principal, nomes_vulgares
#     try:
#         status_label.config(text="Carregando inventário principal...")
#         status_label.pack(pady=10)

#         planilha_principal = pd.read_excel(arquivo1, engine="openpyxl")
#         colunas_existentes = [col for col in planilha_principal.columns if col in ["Nome Vulgar"]]
#         if not colunas_existentes:
#             raise ValueError("A planilha principal não possui a coluna 'Nome Vulgar'.")
#         nomes_vulgares = sorted(planilha_principal["Nome Vulgar"].dropna().unique())
#         atualizar_listbox_nomes("")  # Inicializa a Listbox com todos os nomes

#         frame_listbox.pack(pady=10)
#         frame_secundario.pack(pady=10)
#     except Exception as e:
#         tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha principal: {e}")
#     finally:
#         status_label.pack_forget()


# def carregar_planilha_secundaria(arquivo2):
#     """Carrega a planilha secundária em segundo plano."""
#     global planilha_secundaria
#     try:
#         planilha_secundaria = pd.read_excel(arquivo2, engine="openpyxl")
#         print("Planilha secundária carregada com sucesso.")
#     except Exception as e:
#         tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha secundária: {e}")


# def atualizar_listbox_nomes(filtro):
#     """Atualiza a Listbox com nomes vulgares que atendem ao filtro."""
#     listbox_nomes_vulgares.delete(0, tk.END)
#     for nome in nomes_vulgares:
#         if filtro.lower() in nome.lower():
#             listbox_nomes_vulgares.insert(tk.END, nome)


# def atualizar_listbox_selecionados():
#     """Atualiza a Listbox com os nomes selecionados."""
#     listbox_selecionados.delete(0, tk.END)
#     for nome in nomes_selecionados:
#         listbox_selecionados.insert(tk.END, nome)


# def pesquisar_nomes(event):
#     """Callback para filtrar nomes com base na pesquisa."""
#     filtro = pesquisa_var.get()
#     atualizar_listbox_nomes(filtro)


# def adicionar_selecao(event):
#     """Adiciona um nome à lista de selecionados ao clicar."""
#     global nomes_selecionados

#     selecao = listbox_nomes_vulgares.curselection()
#     if selecao:
#         nome = listbox_nomes_vulgares.get(selecao[0])  # Obtém o nome selecionado
#         if nome not in nomes_selecionados:
#             nomes_selecionados.append(nome)
#             atualizar_listbox_selecionados()


# def remover_ultimo_selecionado():
#     """Remove o último nome adicionado à lista de selecionados."""
#     if nomes_selecionados:
#         nomes_selecionados.pop()
#         atualizar_listbox_selecionados()


# def limpar_lista_selecionados():
#     """Limpa todos os nomes da lista de selecionados."""
#     nomes_selecionados.clear()
#     atualizar_listbox_selecionados()

# def processar_planilhas():
#     """Processa os dados da planilha principal e mescla com nomes científicos."""
#     global planilha_principal

#     arquivo2 = entrada2_var.get()

#     # Verificar se a planilha principal já foi carregada
#     if planilha_principal is None:
#         messagebox.showerror("Erro", "Por favor, selecione e aguarde o carregamento da planilha principal.")
#         return

#     # Verificar se o segundo arquivo foi selecionado
#     if not arquivo2:
#         messagebox.showerror("Erro", "Por favor, selecione o arquivo de Nomes Vulgares e Científicos.")
#         return

#     try:
#         start_time = time.time()
#         progress_var.set(10)  # Atualiza o progresso inicial
#         progress_bar.pack(pady=10, fill="x")  # Mostra a barra de progresso
#         progress_bar.update()

#         # Criar o DataFrame de saída
#         print("Criando o DataFrame de saída...")
#         df_saida = pd.DataFrame()
#         for entrada, saida in {
#             "UT": "UT",
#             "Faixa": "Faixa",
#             "Placa": "Placa",
#             "Nome Vulgar": "Nome Vulgar",
#             "CAP": "CAP",
#             "ALT": "ALT",
#             "QF": "QF",
#             "X": "X",
#             "Y": "Y",
#             "DAP": "DAP",
#             "Volumes (m³)": "Volume_m3",
#             "Latitude": "Latitude",
#             "Longitude": "Longitude",
#             "DM": "DM",
#             "Observações": "Observacoes"
#         }.items():
#             if entrada in planilha_principal.columns:
#                 df_saida[saida] = planilha_principal[entrada]
#             else:
#                 df_saida[saida] = None  # Colunas ausentes preenchidas com None
#         if "Categoria" not in df_saida.columns:
#             df_saida["Categoria"] = None
#         progress_var.set(40)
#         progress_bar.update()

#         # Carregar a segunda planilha
#         print("Carregando a segunda planilha...")
#         df2 = pd.read_excel(arquivo2, engine="openpyxl")
#         df2.rename(columns={
#             "NOME_VULGAR": "Nome Vulgar", 
#             "NOME_CIENTIFICO": "Nome Cientifico",
#             "SITUACAO": "Situacao"
#         }, inplace=True)

#         # Normalizar as colunas "Nome Vulgar" e "Nome Cientifico" e mesclar os dados
#         df_saida["Nome Vulgar"] = df_saida["Nome Vulgar"].str.strip().str.upper()
#         df2["Nome Vulgar"] = df2["Nome Vulgar"].str.strip().str.upper()
#         df2["Nome Cientifico"] = df2["Nome Cientifico"].str.strip().str.upper()
#         df_saida = pd.merge(df_saida, df2[["Nome Vulgar", "Nome Cientifico", "Situacao"]], 
#                             on="Nome Vulgar", how="left")
#         #categorizando as arvores REM
#         # Atualizar a coluna "Categoria" com "REM" se a situação for "protegida"
#         df_saida.loc[df_saida["Situacao"].str.lower() == "protegida", "Categoria"] = "REM"
#         # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'DAP' < 0.5
#         df_saida.loc[df_saida["DAP"] < 0.5, "Categoria"] = "REM"
#         # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'DAP' >= 2
#         df_saida.loc[df_saida["DAP"] >= 2, "Categoria"] = "REM"
#         # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'QF' = 3
#         df_saida.loc[df_saida["QF"] == 3, "Categoria"] = "REM"
#         # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'ALT' > (DEFINITO PELO USUARIO)
#         # Criterio e opcional
#         # df_saida.loc[df_saida["ALT"] >= 0, "Categoria" ] = "REM"



#         # #apagar acoluna "situação após a ultilizacao"
#         # df_saida = df_saida.drop(columns="Situacao")

#         # Reordenar as colunas para garantir que "Nome Cientifico" esteja ao lado de "Nome Vulgar"
#         colunas = list(df_saida.columns)
#         if "Nome Cientifico" in colunas:
#             colunas.remove("Nome Cientifico")
#             idx = colunas.index("Nome Vulgar") + 1
#             colunas.insert(idx, "Nome Cientifico")
#         df_saida = df_saida[colunas]

#         # Salvar o arquivo de saída
#         print("Salvando o arquivo de saída...")
#         diretorio = os.path.dirname(entrada1_var.get())
#         arquivo_saida = os.path.join(diretorio, "planilha_processada_completa.xlsx")
#         df_saida.to_excel(arquivo_saida, index=False, engine="xlsxwriter")
#         progress_var.set(100)
#         progress_bar.update()

#         elapsed_time = time.time() - start_time
#         print(f"Processamento concluído em {elapsed_time:.2f} segundos.")
#         messagebox.showinfo("Sucesso", f"Planilha processada salva em:\n{arquivo_saida}\nTempo total: {elapsed_time:.2f} segundos")
#     except Exception as e:
#         messagebox.showerror("Erro", f"Falha ao processar as planilhas: {e}")
#     finally:
#         progress_var.set(0)  # Resetar a barra de progresso
#         progress_bar.pack_forget()  # Esconde a barra de progresso
#         progress_bar.update()

# def iniciar_processamento():
#     """Inicia o processamento em uma thread separada."""
#     thread = threading.Thread(target=processar_planilhas)
#     thread.daemon = True  # Fecha a thread quando a interface é fechada
#     thread.start()
# # def processar_planilhas():
# #     """Processa os dados das planilhas."""
# #     global planilha_principal, planilha_secundaria

# #     if planilha_principal is None:
# #         tk.messagebox.showerror("Erro", "Por favor, carregue a planilha principal.")
# #         return

# #     if planilha_secundaria is None:
# #         tk.messagebox.showerror("Erro", "Por favor, carregue a planilha secundária.")
# #         return

# #     try:
# #         print("Processando planilhas...")
# #         df_saida = planilha_principal.copy()

# #         # Processar dados da planilha secundária
# #         if "Nome Vulgar" in planilha_secundaria.columns:
# #             planilha_secundaria["Nome Vulgar"] = planilha_secundaria["Nome Vulgar"].str.strip().str.upper()
# #             df_saida = pd.merge(
# #                 df_saida, 
# #                 planilha_secundaria[["Nome Vulgar", "Nome Cientifico"]], 
# #                 on="Nome Vulgar", 
# #                 how="left"
# #             )

# #         # Salvar o arquivo de saída
# #         diretorio = os.path.dirname(entrada1_var.get())
# #         arquivo_saida = os.path.join(diretorio, "planilha_processada_completa.xlsx")
# #         df_saida.to_excel(arquivo_saida, index=False, engine="xlsxwriter")
# #         tk.messagebox.showinfo("Sucesso", f"Processamento concluído! Planilha salva em: {arquivo_saida}")

# #     except Exception as e:
# #         tk.messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")


# # Interface gráfica
# app = tk.Tk()
# app.title("IFDIGITAL 3.0")
# app.geometry("900x700")

# entrada1_var = tk.StringVar()
# entrada2_var = tk.StringVar()
# pesquisa_var = tk.StringVar()

# # Frame para entrada de arquivos
# frame_inputs = ttk.LabelFrame(app, text="Entrada de Arquivos", padding=(10, 10))
# frame_inputs.pack(fill="x", pady=10, padx=10)

# ttk.Label(frame_inputs, text="Arquivo 1: Planilha Principal").grid(row=0, column=0, sticky="w")
# ttk.Entry(frame_inputs, textvariable=entrada1_var, width=60).grid(row=0, column=1, pady=5, padx=5)
# ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("principal")).grid(row=0, column=2, padx=5)

# ttk.Label(frame_inputs, text="Arquivo 2: Planilha Secundária").grid(row=1, column=0, sticky="w")
# ttk.Entry(frame_inputs, textvariable=entrada2_var, width=60).grid(row=1, column=1, pady=5, padx=5)
# ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("secundaria")).grid(row=1, column=2, padx=5)

# status_label = ttk.Label(app, text="")

# # Frame para Listboxes
# frame_listbox = ttk.LabelFrame(app, text="Seleção de Nomes Vulgares", padding=(10, 10))

# ttk.Label(frame_listbox, text="Pesquisar:").grid(row=0, column=0, sticky="w")
# pesquisa_entry = ttk.Entry(frame_listbox, textvariable=pesquisa_var, width=40)
# pesquisa_entry.grid(row=0, column=1, padx=10, pady=5)
# pesquisa_entry.bind("<KeyRelease>", pesquisar_nomes)

# listbox_nomes_vulgares = tk.Listbox(frame_listbox, selectmode=tk.SINGLE, width=40, height=20)
# listbox_nomes_vulgares.bind("<<ListboxSelect>>", adicionar_selecao)
# listbox_nomes_vulgares.grid(row=1, column=0, padx=10, pady=10)

# listbox_selecionados = tk.Listbox(frame_listbox, width=40, height=20)
# listbox_selecionados.grid(row=1, column=1, padx=10, pady=10)

# ttk.Button(frame_listbox, text="Remover Último", command=remover_ultimo_selecionado).grid(row=2, column=0, pady=10)
# ttk.Button(frame_listbox, text="Limpar Lista", command=limpar_lista_selecionados).grid(row=2, column=1, pady=10)

# # Frame para o botão de processamento
# frame_secundario = ttk.Frame(app, padding=(10, 10))
# ttk.Button(frame_secundario, text="Processar Planilhas", command=processar_planilhas, width=40).pack(pady=10)

# frame_listbox.pack_forget()  # Inicialmente escondido
# frame_secundario.pack_forget()  # Inicialmente escondido

# app.mainloop()
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time

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
    "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria"
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
        df2.rename(columns={
            "NOME_VULGAR": "Nome Vulgar", 
            "NOME_CIENTIFICO": "Nome Cientifico",
            "SITUACAO": "Situacao"
        }, inplace=True)

        # Normalizar as colunas "Nome Vulgar" e "Nome Cientifico" e mesclar os dados
        df_saida["Nome Vulgar"] = df_saida["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Vulgar"] = df2["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Cientifico"] = df2["Nome Cientifico"].str.strip().str.upper()
        df_saida = pd.merge(df_saida, df2[["Nome Vulgar", "Nome Cientifico", "Situacao"]], 
                            on="Nome Vulgar", how="left")
        df_saida.loc[df_saida["Situacao"].str.lower() == "protegida", "Categoria"] = "REM"
        df_saida.loc[df_saida["DAP"] < 0.5, "Categoria"] = "REM"
        df_saida.loc[df_saida["DAP"] >= 2, "Categoria"] = "REM"
        df_saida.loc[df_saida["QF"] == 3, "Categoria"] = "REM"

        # Reordenar as colunas
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
        progress_var.set(0)
        progress_bar.pack_forget()
        progress_bar.update()


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

# Barra de progresso
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(app, orient="horizontal", length=400, mode="determinate", variable=progress_var)

# Frame para o botão de processamento
frame_secundario = ttk.Frame(app, padding=(10, 10))
ttk.Button(frame_secundario, text="Processar Planilhas", command=iniciar_processamento, width=40).pack(pady=10)

frame_listbox.pack_forget()  # Inicialmente escondido
frame_secundario.pack_forget()  # Inicialmente escondido

app.mainloop()


# import pandas as pd
# import tkinter as tk
# from tkinter import filedialog, messagebox, ttk
# import threading
# import os
# import time

# # Variáveis globais
# planilha_principal = None
# planilha_secundaria = None
# nomes_vulgares = []  # Lista de todos os nomes vulgares
# nomes_selecionados = []  # Lista para manter a ordem dos nomes selecionados

# # Colunas de entrada e saída
# COLUNAS_ENTRADA = [
#     "Folha", "Secção", "UT", "Faixa", "Placa", "Cod.", "Nome Vulgar", "CAP", "ALT", "QF",
#     "X", "Y", "DAP", "Volumes (m³)", "Latitude", "Longitude", "DM", "Observações"
# ]

# COLUNAS_SAIDA = [
#     "UT", "Faixa", "Placa", "Nome Vulgar",  "Nome Cientifico" , "CAP", "ALT", "QF", "X", "Y",
#     "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria"
# ]

# # Funções da Interface e Processamento

# def selecionar_arquivos(tipo):
#     """Seleciona os arquivos das planilhas."""
#     global planilha_principal, planilha_secundaria

#     arquivo = filedialog.askopenfilename(
#         title=f"Selecione a planilha {'principal' if tipo == 'principal' else 'secundária'}",
#         filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
#     )
#     if arquivo:
#         if tipo == "principal":
#             entrada1_var.set(arquivo)
#             threading.Thread(target=carregar_planilha_principal, args=(arquivo,)).start()
#         elif tipo == "secundaria":
#             entrada2_var.set(arquivo)
#             threading.Thread(target=carregar_planilha_secundaria, args=(arquivo,)).start()


# def carregar_planilha_principal(arquivo1):
#     """Carrega a planilha principal e exibe os nomes vulgares."""
#     global planilha_principal, nomes_vulgares
#     try:
#         status_label.config(text="Carregando inventário principal...")
#         status_label.pack(pady=10)

#         planilha_principal = pd.read_excel(arquivo1, engine="openpyxl")
#         colunas_existentes = [col for col in planilha_principal.columns if col in ["Nome Vulgar"]]
#         if not colunas_existentes:
#             raise ValueError("A planilha principal não possui a coluna 'Nome Vulgar'.")
#         nomes_vulgares = sorted(planilha_principal["Nome Vulgar"].dropna().unique())
#         atualizar_listbox_nomes("")  # Inicializa a Listbox com todos os nomes

#         frame_listbox.pack(pady=10)
#         frame_secundario.pack(pady=10)
#     except Exception as e:
#         tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha principal: {e}")
#     finally:
#         status_label.pack_forget()


# def carregar_planilha_secundaria(arquivo2):
#     """Carrega a planilha secundária em segundo plano."""
#     global planilha_secundaria
#     try:
#         planilha_secundaria = pd.read_excel(arquivo2, engine="openpyxl")
#         print("Planilha secundária carregada com sucesso.")
#     except Exception as e:
#         tk.messagebox.showerror("Erro", f"Erro ao carregar a planilha secundária: {e}")


# def atualizar_listbox_nomes(filtro):
#     """Atualiza a Listbox com nomes vulgares que atendem ao filtro."""
#     listbox_nomes_vulgares.delete(0, tk.END)
#     for nome in nomes_vulgares:
#         if filtro.lower() in nome.lower():
#             listbox_nomes_vulgares.insert(tk.END, nome)


# def atualizar_listbox_selecionados():
#     """Atualiza a Listbox com os nomes selecionados."""
#     listbox_selecionados.delete(0, tk.END)
#     for nome in nomes_selecionados:
#         listbox_selecionados.insert(tk.END, nome)


# def pesquisar_nomes(event):
#     """Callback para filtrar nomes com base na pesquisa."""
#     filtro = pesquisa_var.get()
#     atualizar_listbox_nomes(filtro)


# def adicionar_selecao(event):
#     """Adiciona um nome à lista de selecionados ao clicar."""
#     global nomes_selecionados

#     selecao = listbox_nomes_vulgares.curselection()
#     if selecao:
#         nome = listbox_nomes_vulgares.get(selecao[0])  # Obtém o nome selecionado
#         if nome not in nomes_selecionados:
#             nomes_selecionados.append(nome)
#             atualizar_listbox_selecionados()


# def remover_ultimo_selecionado():
#     """Remove o último nome adicionado à lista de selecionados."""
#     if nomes_selecionados:
#         nomes_selecionados.pop()
#         atualizar_listbox_selecionados()


# def limpar_lista_selecionados():
#     """Limpa todos os nomes da lista de selecionados."""
#     nomes_selecionados.clear()
#     atualizar_listbox_selecionados()

# def processar_planilhas():
#     """Processa os dados da planilha principal e mescla com nomes científicos."""
#     global planilha_principal

#     arquivo2 = entrada2_var.get()

#     # Verificar se a planilha principal já foi carregada
#     if planilha_principal is None:
#         messagebox.showerror("Erro", "Por favor, selecione e aguarde o carregamento da planilha principal.")
#         return

#     # Verificar se o segundo arquivo foi selecionado
#     if not arquivo2:
#         messagebox.showerror("Erro", "Por favor, selecione o arquivo de Nomes Vulgares e Científicos.")
#         return

#     try:
#         start_time = time.time()
#         progress_var.set(10)  # Atualiza o progresso inicial
#         progress_bar.pack(pady=10, fill="x")  # Mostra a barra de progresso
#         progress_bar.update()

#         # Criar o DataFrame de saída
#         print("Criando o DataFrame de saída...")
#         df_saida = pd.DataFrame()
#         for entrada, saida in {
#             "UT": "UT",
#             "Faixa": "Faixa",
#             "Placa": "Placa",
#             "Nome Vulgar": "Nome Vulgar",
#             "CAP": "CAP",
#             "ALT": "ALT",
#             "QF": "QF",
#             "X": "X",
#             "Y": "Y",
#             "DAP": "DAP",
#             "Volumes (m³)": "Volume_m3",
#             "Latitude": "Latitude",
#             "Longitude": "Longitude",
#             "DM": "DM",
#             "Observações": "Observacoes"
#         }.items():
#             if entrada in planilha_principal.columns:
#                 df_saida[saida] = planilha_principal[entrada]
#             else:
#                 df_saida[saida] = None  # Colunas ausentes preenchidas com None
#         if "Categoria" not in df_saida.columns:
#             df_saida["Categoria"] = None
#         progress_var.set(40)
#         progress_bar.update()

#         # Carregar a segunda planilha
#         print("Carregando a segunda planilha...")
#         df2 = pd.read_excel(arquivo2, engine="openpyxl")
#         df2.rename(columns={
#             "NOME_VULGAR": "Nome Vulgar", 
#             "NOME_CIENTIFICO": "Nome Cientifico",
#             "SITUACAO": "Situacao"
#         }, inplace=True)

#         # Normalizar as colunas "Nome Vulgar" e "Nome Cientifico" e mesclar os dados
#         df_saida["Nome Vulgar"] = df_saida["Nome Vulgar"].str.strip().str.upper()
#         df2["Nome Vulgar"] = df2["Nome Vulgar"].str.strip().str.upper()
#         df2["Nome Cientifico"] = df2["Nome Cientifico"].str.strip().str.upper()
#         df_saida = pd.merge(df_saida, df2[["Nome Vulgar", "Nome Cientifico", "Situacao"]], 
#                             on="Nome Vulgar", how="left")
#         #categorizando as arvores REM
#         # Atualizar a coluna "Categoria" com "REM" se a situação for "protegida"
#         df_saida.loc[df_saida["Situacao"].str.lower() == "protegida", "Categoria"] = "REM"
#         # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'DAP' < 0.5
#         df_saida.loc[df_saida["DAP"] < 0.5, "Categoria"] = "REM"
#         # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'DAP' >= 2
#         df_saida.loc[df_saida["DAP"] >= 2, "Categoria"] = "REM"
#         # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'QF' = 3
#         df_saida.loc[df_saida["QF"] == 3, "Categoria"] = "REM"
#         # Atualiza a coluna 'Categoria' com "REM" para linhas onde 'ALT' > (DEFINITO PELO USUARIO)
#         # Criterio e opcional
#         # df_saida.loc[df_saida["ALT"] >= 0, "Categoria" ] = "REM"



#         # #apagar acoluna "situação após a ultilizacao"
#         # df_saida = df_saida.drop(columns="Situacao")

#         # Reordenar as colunas para garantir que "Nome Cientifico" esteja ao lado de "Nome Vulgar"
#         colunas = list(df_saida.columns)
#         if "Nome Cientifico" in colunas:
#             colunas.remove("Nome Cientifico")
#             idx = colunas.index("Nome Vulgar") + 1
#             colunas.insert(idx, "Nome Cientifico")
#         df_saida = df_saida[colunas]

#         # Salvar o arquivo de saída
#         print("Salvando o arquivo de saída...")
#         diretorio = os.path.dirname(entrada1_var.get())
#         arquivo_saida = os.path.join(diretorio, "planilha_processada_completa.xlsx")
#         df_saida.to_excel(arquivo_saida, index=False, engine="xlsxwriter")
#         progress_var.set(100)
#         progress_bar.update()

#         elapsed_time = time.time() - start_time
#         print(f"Processamento concluído em {elapsed_time:.2f} segundos.")
#         messagebox.showinfo("Sucesso", f"Planilha processada salva em:\n{arquivo_saida}\nTempo total: {elapsed_time:.2f} segundos")
#     except Exception as e:
#         messagebox.showerror("Erro", f"Falha ao processar as planilhas: {e}")
#     finally:
#         progress_var.set(0)  # Resetar a barra de progresso
#         progress_bar.pack_forget()  # Esconde a barra de progresso
#         progress_bar.update()

# def iniciar_processamento():
#     """Inicia o processamento em uma thread separada."""
#     thread = threading.Thread(target=processar_planilhas)
#     thread.daemon = True  # Fecha a thread quando a interface é fechada
#     thread.start()
# # def processar_planilhas():
# #     """Processa os dados das planilhas."""
# #     global planilha_principal, planilha_secundaria

# #     if planilha_principal is None:
# #         tk.messagebox.showerror("Erro", "Por favor, carregue a planilha principal.")
# #         return

# #     if planilha_secundaria is None:
# #         tk.messagebox.showerror("Erro", "Por favor, carregue a planilha secundária.")
# #         return

# #     try:
# #         print("Processando planilhas...")
# #         df_saida = planilha_principal.copy()

# #         # Processar dados da planilha secundária
# #         if "Nome Vulgar" in planilha_secundaria.columns:
# #             planilha_secundaria["Nome Vulgar"] = planilha_secundaria["Nome Vulgar"].str.strip().str.upper()
# #             df_saida = pd.merge(
# #                 df_saida, 
# #                 planilha_secundaria[["Nome Vulgar", "Nome Cientifico"]], 
# #                 on="Nome Vulgar", 
# #                 how="left"
# #             )

# #         # Salvar o arquivo de saída
# #         diretorio = os.path.dirname(entrada1_var.get())
# #         arquivo_saida = os.path.join(diretorio, "planilha_processada_completa.xlsx")
# #         df_saida.to_excel(arquivo_saida, index=False, engine="xlsxwriter")
# #         tk.messagebox.showinfo("Sucesso", f"Processamento concluído! Planilha salva em: {arquivo_saida}")

# #     except Exception as e:
# #         tk.messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")


# # Interface gráfica
# app = tk.Tk()
# app.title("IFDIGITAL 3.0")
# app.geometry("900x700")

# entrada1_var = tk.StringVar()
# entrada2_var = tk.StringVar()
# pesquisa_var = tk.StringVar()

# # Frame para entrada de arquivos
# frame_inputs = ttk.LabelFrame(app, text="Entrada de Arquivos", padding=(10, 10))
# frame_inputs.pack(fill="x", pady=10, padx=10)

# ttk.Label(frame_inputs, text="Arquivo 1: Planilha Principal").grid(row=0, column=0, sticky="w")
# ttk.Entry(frame_inputs, textvariable=entrada1_var, width=60).grid(row=0, column=1, pady=5, padx=5)
# ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("principal")).grid(row=0, column=2, padx=5)

# ttk.Label(frame_inputs, text="Arquivo 2: Planilha Secundária").grid(row=1, column=0, sticky="w")
# ttk.Entry(frame_inputs, textvariable=entrada2_var, width=60).grid(row=1, column=1, pady=5, padx=5)
# ttk.Button(frame_inputs, text="Selecionar", command=lambda: selecionar_arquivos("secundaria")).grid(row=1, column=2, padx=5)

# status_label = ttk.Label(app, text="")

# # Frame para Listboxes
# frame_listbox = ttk.LabelFrame(app, text="Seleção de Nomes Vulgares", padding=(10, 10))

# ttk.Label(frame_listbox, text="Pesquisar:").grid(row=0, column=0, sticky="w")
# pesquisa_entry = ttk.Entry(frame_listbox, textvariable=pesquisa_var, width=40)
# pesquisa_entry.grid(row=0, column=1, padx=10, pady=5)
# pesquisa_entry.bind("<KeyRelease>", pesquisar_nomes)

# listbox_nomes_vulgares = tk.Listbox(frame_listbox, selectmode=tk.SINGLE, width=40, height=20)
# listbox_nomes_vulgares.bind("<<ListboxSelect>>", adicionar_selecao)
# listbox_nomes_vulgares.grid(row=1, column=0, padx=10, pady=10)

# listbox_selecionados = tk.Listbox(frame_listbox, width=40, height=20)
# listbox_selecionados.grid(row=1, column=1, padx=10, pady=10)

# ttk.Button(frame_listbox, text="Remover Último", command=remover_ultimo_selecionado).grid(row=2, column=0, pady=10)
# ttk.Button(frame_listbox, text="Limpar Lista", command=limpar_lista_selecionados).grid(row=2, column=1, pady=10)

# # Frame para o botão de processamento
# frame_secundario = ttk.Frame(app, padding=(10, 10))
# ttk.Button(frame_secundario, text="Processar Planilhas", command=processar_planilhas, width=40).pack(pady=10)

# frame_listbox.pack_forget()  # Inicialmente escondido
# frame_secundario.pack_forget()  # Inicialmente escondido

# app.mainloop()
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time

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
    "DAP", "Volume_m3", "Latitude", "Longitude", "DM", "Observacoes", "Categoria"
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
        df2.rename(columns={
            "NOME_VULGAR": "Nome Vulgar", 
            "NOME_CIENTIFICO": "Nome Cientifico",
            "SITUACAO": "Situacao"
        }, inplace=True)

        # Normalizar as colunas "Nome Vulgar" e "Nome Cientifico" e mesclar os dados
        df_saida["Nome Vulgar"] = df_saida["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Vulgar"] = df2["Nome Vulgar"].str.strip().str.upper()
        df2["Nome Cientifico"] = df2["Nome Cientifico"].str.strip().str.upper()
        df_saida = pd.merge(df_saida, df2[["Nome Vulgar", "Nome Cientifico", "Situacao"]], 
                            on="Nome Vulgar", how="left")
        df_saida.loc[df_saida["Situacao"].str.lower() == "protegida", "Categoria"] = "REM"
        df_saida.loc[df_saida["DAP"] < 0.5, "Categoria"] = "REM"
        df_saida.loc[df_saida["DAP"] >= 2, "Categoria"] = "REM"
        df_saida.loc[df_saida["QF"] == 3, "Categoria"] = "REM"

        # Reordenar as colunas
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
        progress_var.set(0)
        progress_bar.pack_forget()
        progress_bar.update()


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

# Barra de progresso
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(app, orient="horizontal", length=400, mode="determinate", variable=progress_var)

# Frame para o botão de processamento
frame_secundario = ttk.Frame(app, padding=(10, 10))
ttk.Button(frame_secundario, text="Processar Planilhas", command=iniciar_processamento, width=40).pack(pady=10)

frame_listbox.pack_forget()  # Inicialmente escondido
frame_secundario.pack_forget()  # Inicialmente escondido

app.mainloop()
