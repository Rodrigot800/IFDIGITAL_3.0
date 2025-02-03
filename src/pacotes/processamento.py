import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import time
from edicaoValorFiltro import  abrir_janela_valores_padroes,valor1,valor2,valor3,valor4
import os
import configparser

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
        print(df_saida[["UT", "UT_ID", "UT_AREA_HA"]].drop_duplicates())
        
        #indice de raridade e classificação de substituta 
        ####


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

