import tkinter as tk
from tkinter import ttk

# Função para adicionar dados com alternância de cor
def inserir_dados():
    # Dados fictícios
    data = [
        ("UT1", "10", "50", "100", "200", "30", "10", "150", "20%", "10", "5"),
        ("UT2", "12", "60", "120", "220", "32", "12", "160", "22%", "12", "6"),
        ("UT3", "14", "70", "130", "230", "34", "14", "170", "24%", "14", "7"),
        ("UT4", "16", "80", "140", "240", "36", "16", "180", "26%", "16", "8")
    ]
    
    # Alternar cores (branco e verde claro)
    for i, row in enumerate(data):
        if i % 2 == 0:  # Linhas com índice par
            table_ut_vol.insert("", "end", values=row, tags=('verde_claro',))
        else:  # Linhas com índice ímpar
            table_ut_vol.insert("", "end", values=row, tags=('branca',))
            

# Criar janela principal
root = tk.Tk()
root.title("Alternância de Cor no Treeview")

# Criar a tabela Treeview
table_ut_vol = ttk.Treeview(root, columns=("UT", "Hectares", "n° Árv", "Vol(m³)", "Vol_Max", "Diminuir", "Aumentar", "V_m³/ha", "DAP %", "CAP", "ALT"), show="headings")

# Configurar as colunas
colunas = ["UT", "Hectares", "n° Árv", "Vol(m³)", "Vol_Max", "Diminuir", "Aumentar", "V_m³/ha", "DAP %", "CAP", "ALT"]
for col in colunas:
    table_ut_vol.heading(col, text=col)
    table_ut_vol.column(col, width=100, anchor="center")

# Definindo as tags para alternar cores
table_ut_vol.tag_configure('branca', background='white')  # Cor para as linhas brancas
table_ut_vol.tag_configure('verde_claro', background='#d3f8e2')  # Cor para as linhas verde claro

# Inserir dados com alternância de cores
inserir_dados()

# Mostrar a tabela
table_ut_vol.pack(pady=20)

# Iniciar o loop da interface gráfica
root.mainloop()
