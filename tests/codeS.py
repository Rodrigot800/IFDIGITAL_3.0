import tkinter as tk
from tkinter import ttk

# Suponha que 'frame_listbox' já esteja criado e posicionado
# Por exemplo:
app = tk.Tk()
app.title("Exemplo com Treeview")
frame_listbox = ttk.Frame(app)
frame_listbox.pack(padx=10, pady=10, fill="both", expand=True)

# Defina as colunas para a tabela de selecionados
colunas_selecionados = ("Nome", "DAP <", "DAP >=", "QF", "ALT >", "CAP <")

# Cria o Treeview que atuará como tabela
table_selecionados = ttk.Treeview(frame_listbox, columns=colunas_selecionados, show="headings", height=20)
for col in colunas_selecionados:
    table_selecionados.heading(col, text=col)
    table_selecionados.column(col, width=100, anchor="center")
table_selecionados.grid(row=1, column=1, padx=10, pady=10)

# Exemplo de widget de onde virão as seleções (pode ser uma listbox ou outro)
# Para este exemplo, criamos uma Listbox com alguns nomes:
listbox_nomes = tk.Listbox(frame_listbox, selectmode=tk.SINGLE, width=40, height=10)
nomes_exemplo = ["Nome1", "Nome2", "Nome3", "Nome4"]
for nome in nomes_exemplo:
    listbox_nomes.insert(tk.END, nome)
listbox_nomes.grid(row=1, column=0, padx=10, pady=10)

def adicionar_selecao(event):
    # Obtém os índices selecionados na Listbox
    indices = listbox_nomes.curselection()
    if not indices:
        return
    index = indices[0]
    # Obtém o nome selecionado
    nome = listbox_nomes.get(index)
    
    # Verifica se o item já foi adicionado à tabela (Treeview)
    ja_adicionado = False
    for child in table_selecionados.get_children():
        # Supondo que o nome esteja na primeira coluna
        valores = table_selecionados.item(child, "values")
        if valores and valores[0] == nome:
            ja_adicionado = True
            break
    
    # Se o item ainda não estiver na tabela, insere-o
    if not ja_adicionado:
        default_values = (nome, 0.5, 2, 3, "", "")
        table_selecionados.insert("", "end", values=default_values)


def selecionar_todos():
    """Adiciona todos os itens da listbox_nomes à tabela, evitando duplicatas."""
    # Itera por todos os índices da listbox
    for i in range(listbox_nomes.size()):
        nome = listbox_nomes.get(i)
        # Verifica se o item já foi adicionado à tabela (Treeview)
        ja_adicionado = False
        for child in table_selecionados.get_children():
            valores = table_selecionados.item(child, "values")
            if valores and valores[0] == nome:
                ja_adicionado = True
                break
        # Se não estiver presente, insere com valores padrão para as demais colunas
        if not ja_adicionado:
            default_values = (nome, 0.5, 2, 3, "", "")
            table_selecionados.insert("", "end", values=default_values)



# Função para remover o último item da tabela
def remover_ultimo_selecionado():
    filhos = table_selecionados.get_children()
    if filhos:
        ultimo = filhos[-1]
        table_selecionados.delete(ultimo)

# Função para limpar todos os itens da tabela
def limpar_lista_selecionados():
    for item in table_selecionados.get_children():
        table_selecionados.delete(item)

# Vincula o evento de seleção (por exemplo, clique duplo) para adicionar à tabela
listbox_nomes.bind("<<ListboxSelect>>", adicionar_selecao)

# Botões para testar as funções de remoção e limpeza
btn_remover = ttk.Button(frame_listbox, text="Remover Último", command=remover_ultimo_selecionado)
btn_remover.grid(row=2, column=0, padx=5, pady=10)

btn_limpar = ttk.Button(frame_listbox, text="Limpar Lista", command=limpar_lista_selecionados)
btn_limpar.grid(row=2, column=1, padx=5, pady=10)
btn_limpar = ttk.Button(frame_listbox, text="selecioanar tudo", command=selecionar_todos)
btn_limpar.grid(row=2, column=2, padx=5, pady=10)


app.mainloop()
