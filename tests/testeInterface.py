import tkinter as tk
from tkinter import ttk

# Função para adicionar seleção do Listbox à Tabela
def adicionar_selecao(event):
    # Obter o item selecionado
    index = listbox_nomes_vulgares.curselection()
    if not index:
        return
    nome = listbox_nomes_vulgares.get(index[0])
    table_selecionados.insert("", "end", values=(nome, "Valor 1", "Valor 2", "Valor 3"))

# Criando a janela principal
app = tk.Tk()
app.title("Exemplo Listbox e Tabela Lado a Lado")
app.geometry("1000x600")  # Definindo o tamanho da janela

# Frame para os widgets lado a lado
frame_lado_a_lado = tk.Frame(app)
frame_lado_a_lado.pack(padx=10, pady=10, fill="both", expand=True)

# Criando o Listbox (lado esquerdo)
frame_listbox = tk.Frame(frame_lado_a_lado)
frame_listbox.pack(side="left", padx=10, pady=10)

listbox_nomes_vulgares = tk.Listbox(frame_listbox, selectmode=tk.SINGLE, width=25, height=10)
listbox_nomes_vulgares.pack(padx=5, pady=5)

# Adicionando alguns itens de exemplo
for nome in ["Espécie 1", "Espécie 2", "Espécie 3"]:
    listbox_nomes_vulgares.insert(tk.END, nome)

# Vinculando o clique na Listbox para adicionar à tabela
listbox_nomes_vulgares.bind("<ButtonRelease-1>", adicionar_selecao)

# Criando a Tabela (lado direito)
frame_tabela = tk.Frame(frame_lado_a_lado)
frame_tabela.pack(side="left", padx=10, pady=10)

colunas_selecionados = ("Nome", "DAP <", "DAP >=", "QF = ", "ALT >", "CAP <")
table_selecionados = ttk.Treeview(frame_tabela, columns=colunas_selecionados, show="headings", height=10)

# Configurando as colunas
for col in colunas_selecionados:
    table_selecionados.heading(col, text=col)
    table_selecionados.column(col, width=50, anchor="center")
    table_selecionados.column("Nome", width=150, anchor="w")

table_selecionados.pack(padx=5, pady=5)

# Iniciando a interface
app.mainloop()
