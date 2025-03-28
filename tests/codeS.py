import pandas as pd
import tkinter as tk
from tkinter import ttk

# Simulação do DataFrame (para testes)
df_tabelaDeAjusteVol = pd.DataFrame({
    "ut": [1, 2, 3],
    "CAP": [10.0, 20.0, 30.0],
    "ALT": [15.0, 25.0, 35.0]
})

# Lista de colunas editáveis
colunas_editaveis = ["CAP", "ALT"]

# Criando interface Tkinter
root = tk.Tk()
table_ut_vol = ttk.Treeview(root, columns=("ut", "CAP", "ALT"), show="headings")

# Definição das colunas na Treeview
for col in ("ut", "CAP", "ALT"):
    table_ut_vol.heading(col, text=col)
    table_ut_vol.column(col, width=100)

# Preenchendo a tabela com dados do DataFrame
for index, row in df_tabelaDeAjusteVol.iterrows():
    table_ut_vol.insert("", "end", values=(row["ut"], row["CAP"], row["ALT"]))

table_ut_vol.pack()

def editar_celula_volume(event):
    """Permite editar apenas as colunas CAP e ALT ao dar duplo clique."""
    global entry, item_selecionado, col_index, col_nome

    # Obtém o item e a coluna clicada
    item_selecionado = table_ut_vol.focus()  # Captura a linha selecionada
    coluna_selecionada = table_ut_vol.identify_column(event.x)

    if not item_selecionado:
        return

    # Obtém índice da coluna (converte "#x" para índice, ex: "#2" → 1)
    col_index = int(coluna_selecionada[1:]) - 1
    col_nome = table_ut_vol["columns"][col_index]  # Nome da coluna

    # Verifica se a coluna é editável
    if col_nome not in colunas_editaveis:
        return

    # Obtém as coordenadas da célula
    x, y, largura, altura = table_ut_vol.bbox(item_selecionado, col_index)

    # Obtém o valor atual
    valores = table_ut_vol.item(item_selecionado, "values")
    valor_atual = valores[col_index]

    # Criar Entry para edição
    entry = tk.Entry(table_ut_vol)
    entry.place(x=x, y=y, width=largura, height=altura)
    entry.insert(0, valor_atual)
    entry.focus()

    # Função para salvar e remover o Entry
    def salvar_novo_valor(event=None):
        global df_tabelaDeAjusteVol  # Garante acesso ao DataFrame global
        novo_valor = entry.get()

        try:
            novo_valor = float(novo_valor)  # Converte para número
        except ValueError:
            entry.destroy()
            return  # Sai sem salvar se o valor não for numérico

        # Atualiza a exibição na Treeview
        valores_atualizados = list(table_ut_vol.item(item_selecionado, "values"))
        valores_atualizados[col_index] = novo_valor
        table_ut_vol.item(item_selecionado, values=valores_atualizados)

        # Obtém o valor da coluna "ut" para encontrar a linha correspondente no DataFrame
        ut_val = float(valores_atualizados[0])  # O primeiro valor da linha é o "ut"

        # Verifica se a coluna "ut" existe antes de acessar
        if "ut" not in df_tabelaDeAjusteVol.columns:
            print("Erro: A coluna 'ut' não existe no DataFrame!")
            entry.destroy()
            return

        # Encontra o índice correto no DataFrame
        index_df = df_tabelaDeAjusteVol[df_tabelaDeAjusteVol["ut"] == ut_val].index

        if not index_df.empty:
            df_tabelaDeAjusteVol.loc[index_df, col_nome] = novo_valor  # Atualiza no DataFrame
            print(f"Valor atualizado: UT={ut_val}, {col_nome}={novo_valor}")  # Debug
        else:
            print(f"Erro: UT {ut_val} não encontrado no DataFrame!")

        entry.destroy()

    # Bind para salvar ao pressionar Enter ou sair do campo
    entry.bind("<Return>", salvar_novo_valor)
    entry.bind("<FocusOut>", salvar_novo_valor)

# Associa evento de clique duplo à função de edição
table_ut_vol.bind("<Double-1>", editar_celula_volume)

root.mainloop() 
