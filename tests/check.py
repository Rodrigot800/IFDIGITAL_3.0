histpyinstaller --onefile --windowed --hidden-import=openpyxl --hidden-import=pandas --hidden-import=tkinter --hidden-import=configparser --hidden-import=numpy --hidden-import=xlsxwriter --collect-all openpyxl --collect-all pandas --collect-all numpy --collect-all xlsxwriter --icon="icone.ico" main.py



##### Classificar Substitutas 
        # **Filtrar apenas as árvores que estão como CORTE e com Nome Vulgar nos selecionados**
        df_filtrado = df_saida[
            (df_saida["Categoria"] == "CORTE") & 
            (df_saida["Nome Vulgar"].isin(nomes_selecionados))
        ].copy()

        # **Ordenar: por UT, QF (maior para menor) e Volume_m3 (menor para maior)**
        df_filtrado.sort_values(by=["UT", "QF", "Volume_m3"], ascending=[True, False, True], inplace=True)

        # **Garantir que df_contagem tenha apenas uma linha por UT e Nome Vulgar**
        df_contagem_agg = df_contagem.groupby(["UT", "Nome Vulgar"], as_index=False).agg({"Valor_Substituta": "sum"})

        # **Mesclar com df_contagem_agg para garantir que a quantidade de substitutas seja específica para cada UT e Nome Vulgar**
        df_filtrado = df_filtrado.merge(
            df_contagem_agg,  # Usamos a versão agregada da contagem
            on=["UT", "Nome Vulgar"],
            how="left"
        )

        # Exibir os primeiros resultados para validar a junção
        print("\n--- Validação: df_filtrado após o merge ---")
        print(df_filtrado.head())

        # **Função para definir as árvores substitutas corretamente**
        def definir_substituta(df):
            df["Marcador"] = False  # Criar coluna auxiliar para identificar as árvores que serão substituídas

            # Iterar por UT e Nome Vulgar
            for (ut, nome), grupo in df.groupby(["UT", "Nome Vulgar"]):
                quantidade_substituir = grupo["Valor_Substituta"].iloc[0]  # Obter a quantidade correta para esta UT e Nome Vulgar

                if pd.notna(quantidade_substituir) and quantidade_substituir > 0:
                    indices_para_substituir = grupo.index[:int(quantidade_substituir)]  # Selecionar os primeiros X indivíduos para substituição
                    df.loc[indices_para_substituir, "Marcador"] = True  # Marcar os que devem ser substituídos

            # Aplicar a substituição apenas para os marcados
            df.loc[df["Marcador"], "Categoria"] = "SUBSTITUTA"
            df.drop(columns=["Marcador"], inplace=True)  # Remover a coluna auxiliar

            return df

        # **Aplicar a função para categorizar corretamente como SUBSTITUTA**
        df_filtrado = definir_substituta(df_filtrado)

        # **Filtrar apenas os registros que realmente foram substituídos**
        df_substituta = df_filtrado[df_filtrado["Categoria"] == "SUBSTITUTA"][["UT", "Nome Vulgar", "Categoria"]]

        # **Verificar a saída corrigida**
        print("\n--- df_substituta (Apenas os registros que devem ser SUBSTITUTA) ---")
        print(df_substituta.drop_duplicates())

        # **Atualizar SOMENTE os registros corretos em df_saida**
        # Criar índice com UT, Nome Vulgar e uma chave única (Faixa, Placa) para garantir que apenas os corretos sejam substituídos
        df_saida.set_index(["UT", "Nome Vulgar", "Faixa", "Placa"], inplace=True)
        df_filtrado.set_index(["UT", "Nome Vulgar", "Faixa", "Placa"], inplace=True)

        # **Somente substituir onde há correspondência exata**
        df_saida.update(df_filtrado["Categoria"])

        # **Resetar índice após atualização**
        df_saida.reset_index(inplace=True)