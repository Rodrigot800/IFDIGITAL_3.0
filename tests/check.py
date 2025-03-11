import tkinter as tk
from tkinter import ttk

def on_item_double_click(event):
    """Abre uma janela de edição para a linha selecionada."""
    selected_item = tree.focus()
    if not selected_item:
        return

    # Recupera os valores da linha selecionada
    values = tree.item(selected_item, "values")

    # Cria uma nova janela para edição
    edit_window = tk.Toplevel(root)
    edit_window.title("Editar Dados da Linha")
    edit_window.geometry("400x300")

    # Dicionário para guardar os Entry widgets por coluna
    entries = {}

    # Cria um Label e um Entry para cada coluna
    for i, col in enumerate(colunas):
        tk.Label(edit_window, text=col + ":").grid(row=i, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(edit_window, width=30)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entry.insert(0, values[i])
        entries[col] = entry

    # Função para salvar as alterações
    def salvar_alteracoes():
        # Coleta os novos valores dos campos
        new_values = tuple(entries[col].get() for col in colunas)
        # Atualiza o item na Treeview
        tree.item(selected_item, values=new_values)
        edit_window.destroy()

    # Botão para salvar alterações
    tk.Button(edit_window, text="Salvar Alterações", command=salvar_alteracoes).grid(row=len(colunas), column=0, columnspan=2, pady=10)

# Criação da janela principal
root = tk.Tk()
root.title("Tabela com Edição")
root.geometry("800x400")

# Frame para conter o Treeview
frame = ttk.Frame(root)
frame.pack(fill="both", expand=True, padx=10, pady=10)

# Definição das colunas para a tabela
colunas = ("Nome", "DAP <", "DAP", "DAP >=", "QF", "ALT >", "CAP <")

# Cria o Treeview configurado para mostrar apenas os headings
tree = ttk.Treeview(frame, columns=colunas, show="headings")
for col in colunas:
    tree.heading(col, text=col)
    tree.column(col, width=100, anchor="center")

# Insere linhas de exemplo
tree.insert("", "end", values=("Exemplo 1", "0.5", "1.2", "2.3", "3", "0", "4.2"))
tree.insert("", "end", values=("Exemplo 2", "0.7", "1.4", "2.6", "5", "1", "3.8"))
tree.insert("", "end", values=("Exemplo 3", "0.6", "1.3", "2.4", "4", "0", "4.0"))
tree.pack(fill="both", expand=True)

# Vincula o evento de duplo clique à função que abre a janela de edição
tree.bind("<Double-1>", on_item_double_click)

root.mainloop()
