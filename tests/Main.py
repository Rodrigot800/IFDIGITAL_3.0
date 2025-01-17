import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

def selecionar_arquivo():
    """Abre o explorador de arquivos para selecionar a planilha de entrada e analisa os valores da coluna 'animal'."""
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha de entrada",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if arquivo:
        entrada_var.set(arquivo)
        try:
            # Ler a planilha para identificar os valores únicos de "animal"
            df = pd.read_excel(arquivo)
            if 'animal' not in df.columns:
                messagebox.showerror("Erro", "A planilha selecionada não contém a coluna 'animal'.")
                return
            
            # Obter os valores únicos da coluna "animal"
            valores_unicos = df['animal'].dropna().unique()
            
            # Limpar lista anterior e adicionar novos valores
            lista_animais.delete(0, tk.END)
            for valor in valores_unicos:
                lista_animais.insert(tk.END, valor)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar a planilha: {e}")

def processar_planilha():
    """Processa os dados da planilha e cria um novo arquivo com as modificações."""
    arquivo_entrada = entrada_var.get()
    operacao = operacao_var.get()

    if not arquivo_entrada:
        messagebox.showerror("Erro", "Por favor, selecione o arquivo de entrada.")
        return

    if operacao == "Selecione a operação":
        messagebox.showerror("Erro", "Por favor, selecione uma operação.")
        return

    try:
        # Lendo a planilha de entrada
        df = pd.read_excel(arquivo_entrada)

        # Verificar se as colunas necessárias existem
        if 'animal' not in df.columns or 'altura' not in df.columns or 'peso' not in df.columns:
            messagebox.showerror("Erro", "A planilha deve conter as colunas 'animal', 'altura' e 'peso'.")
            return

        # Filtrar os dados pelos animais selecionados
        indices_selecionados = lista_animais.curselection()
        animais_selecionados = [lista_animais.get(i) for i in indices_selecionados]
        
        if not animais_selecionados:
            messagebox.showerror("Erro", "Por favor, selecione pelo menos um animal para filtrar.")
            return

        df_filtrado = df[df['animal'].isin(animais_selecionados)]

        if df_filtrado.empty:
            messagebox.showerror("Erro", "Nenhum dado corresponde aos filtros selecionados.")
            return

        # Processar a operação escolhida
        if operacao == "Quadrado":
            df_filtrado['peso'] = df_filtrado['peso'] ** 2
        elif operacao == "Cubo":
            df_filtrado['peso'] = df_filtrado['peso'] ** 3

        # Criar um novo arquivo no mesmo diretório
        diretorio = os.path.dirname(arquivo_entrada)
        arquivo_saida = os.path.join(diretorio, "planilha_processada_filtrada.xlsx")

        # Salvando a planilha processada
        df_filtrado.to_excel(arquivo_saida, index=False)

        messagebox.showinfo("Sucesso", f"Planilha processada salva em:\n{arquivo_saida}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao processar a planilha: {e}")

# Criação da interface principal
app = tk.Tk()
app.title("IFDIGITAL 3.0")
app.geometry("500x600")

# Variáveis para os arquivos e a operação
entrada_var = tk.StringVar()
operacao_var = tk.StringVar(value="Selecione a operação")

# Rótulos e botões
tk.Label(app, text="Arquivo de entrada:").pack(pady=5)
tk.Entry(app, textvariable=entrada_var, width=50).pack(pady=5)
tk.Button(app, text="Selecionar Planilha", command=selecionar_arquivo).pack(pady=5)

# Dropdown (ComboBox) para escolher a operação
tk.Label(app, text="Escolha a operação:").pack(pady=5)
operacao_menu = ttk.Combobox(app, textvariable=operacao_var, values=["Selecione a operação", "Quadrado", "Cubo"], state="readonly")
operacao_menu.pack(pady=5)

# Listbox para seleção múltipla de animais
tk.Label(app, text="Filtrar por animal (Selecione múltiplos):").pack(pady=10)
lista_animais = tk.Listbox(app, selectmode=tk.MULTIPLE, height=10, width=40)
lista_animais.pack(pady=5)

# Botão para processar a planilha
tk.Button(app, text="Processar Planilha", command=processar_planilha, bg="green", fg="white").pack(pady=20)

# Rodar a interface
app.mainloop()
