import tkinter as tk
from tkinter import ttk
import pandas as pd

# Criando DataFrame inicial (simulação)
dados = {
    "UT": [101, 102],
    "Hectares": [10.5, 8.3],
    "n° Árv": [25, 18],
    "Vol(m³)": [30.2, 25.1],
    "Vol_Max": [315, 249],
    "Diminuir": [0, 0],
    "Aumentar": [5.2, 8.9],
    "V_m³/ha": [3.2, 3.0],
    "DAP": [20, 19],
    "CAP": [62, 58],
    "ALT": [15, 14]
}
df_modificado = pd.DataFrame(dados)

# Criando interface Tkinter
root = tk.Tk()
root.title("Tabela Editável")

frame_tabela2 = tk.Frame(root)
frame_tabela2.pack(pady=20)

# Definição das colunas
colunas_tabela2 = list(df_modificado.columns)
table_ut_vol = ttk.Treeview(frame_tabela2, columns=colunas_tabela2, show="headings", height=5)

# Configuração das colunas
for col in colunas_tabela2:
    table_ut_vol.heading(col, text=col)
    table_ut_vol.column(col, width=80, anchor="center")

table_ut_vol.pack(pady=10)
table_ut_vol.config(height=10)

# Índices das colunas que podem ser editadas
colunas_editaveis = ["DAP", "CAP", "ALT"]

# Função para editar somente as colunas permitidas
def editar_celula(event):
    """Permite editar apenas as colunas DAP, CAP e ALT ao dar duplo clique."""
    item_selecionado = table_ut_vol.identify_row(event.y)
    coluna_selecionada = table_ut_vol.identify_column(event.x)

    if not item_selecionado:
        return

    # Obtém índice da coluna
    col_index = int(coluna_selecionada[1:]) - 1  # Converte "#x" para índice (ex: "#9" → 8)
    col_nome = colunas_tabela2[col_index]  # Nome da coluna clicada

    # Verifica se a coluna é editável
    if col_nome not in colunas_editaveis:
        return

    # Obtém as coordenadas da célula
    x, y, largura, altura = table_ut_vol.bbox(item_selecionado, col_index)

    # Obtém o valor atual
    valor_atual = table_ut_vol.item(item_selecionado, "values")[col_index]

    # Criar Entry para edição
    entry = tk.Entry(table_ut_vol)
    entry.place(x=x, y=y, width=largura, height=altura)
    entry.insert(0, valor_atual)
    entry.focus()

    # Função para salvar e remover o Entry
    def salvar_novo_valor(event=None):
        novo_valor = entry.get()

        # Verifica se a coluna editada precisa ser um número
        if col_nome in ["DAP", "CAP", "ALT"]:
            try:
                novo_valor = float(novo_valor) if "." in novo_valor else int(novo_valor)
            except ValueError:
                print(f"Erro: '{novo_valor}' não é um número válido.")
                entry.destroy()
                return

        valores_atualizados = list(table_ut_vol.item(item_selecionado, "values"))
        valores_atualizados[col_index] = novo_valor
        table_ut_vol.item(item_selecionado, values=valores_atualizados)
        entry.destroy()

        # Atualiza o DataFrame corretamente
        index_df = int(item_selecionado)  # ID do item na Treeview equivale ao índice no DataFrame
        df_modificado.at[index_df, col_nome] = novo_valor

        print(df_modificado)


    # Bind para salvar ao pressionar Enter ou sair do campo
    entry.bind("<Return>", salvar_novo_valor)
    entry.bind("<FocusOut>", salvar_novo_valor)

# Adiciona evento de duplo clique apenas para colunas específicas
table_ut_vol.bind("<Double-1>", editar_celula)

# Função para preencher a tabela com os dados do DataFrame
def preencher_tabela():
    for i, row in df_modificado.iterrows():
        table_ut_vol.insert("", "end", iid=i, values=list(row))

# Função para salvar o DataFrame modificado
def salvar_dataframe():
    df_modificado.to_excel("dados_editados.xlsx", index=False)
    print("Dados salvos em 'dados_editados.xlsx'")

# Botão para salvar
btn_salvar = tk.Button(root, text="Salvar Dados", command=salvar_dataframe)
btn_salvar.pack(pady=10)

# Preencher a tabela com os dados iniciais
preencher_tabela()

root.mainloop()
