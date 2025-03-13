import tkinter as tk
from tkinter import ttk, messagebox

# Lista global para armazenar os dados (um dicionário para cada nome selecionado)
selected_data_list = []

# Definição das colunas que serão exibidas na tabela
columns = ("Nome", "DAP <", "DAP", "DAP >=", "QF", "ALT >", "CAP <")
def on_item_double_click(event):
    """
    Ao dar duplo clique em uma linha, abre uma janela de edição.
    Os valores da linha serão carregados em campos para edição.
    Ao salvar, os novos valores são armazenados em um dicionário que é adicionado à lista global.
    """
    selected_item = tree.focus()
    if not selected_item:
        return

    # Recupera os valores da linha selecionada
    values = tree.item(selected_item, "values")
    
    # Cria a janela de edição
    edit_window = tk.Toplevel(root)
    edit_window.title("Editar Dados da Linha")
    edit_window.geometry("400x350")
    
    # Dicionário para guardar os widgets Entry para cada coluna
    entries = {}
    for i, col in enumerate(columns):
        tk.Label(edit_window, text=col + ":").grid(row=i, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(edit_window, width=30)
        entry.grid(row=i, column=1, padx=10, pady=5)
        # Preenche o campo com o valor atual da linha
        entry.insert(0, values[i])
        entries[col] = entry

    def salvar_alteracoes():
        # Coleta os novos valores de cada campo em um dicionário
        new_data = {col: entries[col].get() for col in columns}
        # Adiciona os dados à lista global
        selected_data_list.append(new_data)
        messagebox.showinfo("Salvo", "Dados salvos para o nome selecionado.")
        edit_window.destroy()
        # Para visualização, imprime os dados salvos no console
        print("Dados salvos:")
        for data in selected_data_list:
            print(data)
    
    # Botão para salvar as alterações
    tk.Button(edit_window, text="Salvar Alterações", command=salvar_alteracoes).grid(row=len(columns), column=0, columnspan=2, pady=10)