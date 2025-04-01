import tkinter as tk
from tkinter import ttk

# Criando a janela principal
root = tk.Tk()

# Criando os frames
frame_listbox = ttk.Frame(root)
frame_tabela = ttk.Frame(root)

# Usando pack para garantir que os dois frames tenham o mesmo tamanho
frame_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
frame_tabela.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

# Criando uma Listbox dentro do frame_listbox
listbox = tk.Listbox(frame_listbox)
listbox.pack(fill=tk.BOTH, expand=True)

# Criando uma Tabela (ou Treeview) dentro do frame_tabela
treeview = ttk.Treeview(frame_tabela)
treeview.pack(fill=tk.BOTH, expand=True)

# Iniciando a aplicação
root.mainloop()
