# interface.py
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
import threading

# Variáveis globais
planilha_principal = None
planilha_secundaria = None
nomes_vulgares = []  # Lista de todos os nomes vulgares
nomes_selecionados = []  # Lista para manter a ordem dos nomes selecionados


class InterfaceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Processador de Inventário e Mesclagem")
        self.root.geometry("900x700")

        self.entrada1_var = tk.StringVar()
        self.entrada2_var = tk.StringVar()
        self.pesquisa_var = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        """Cria os widgets da interface gráfica."""
        # Frame para entrada de arquivos
        frame_inputs = ttk.LabelFrame(self.root, text="Entrada de Arquivos", padding=(10, 10))
        frame_inputs.pack(fill="x", pady=10, padx=10)

        ttk.Label(frame_inputs, text="Arquivo 1: Planilha Principal").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_inputs, textvariable=self.entrada1_var, width=60).grid(row=0, column=1, pady=5, padx=5)
        ttk.Button(frame_inputs, text="Selecionar", command=lambda: self.selecionar_arquivos("principal")).grid(row=0, column=2, padx=5)

        ttk.Label(frame_inputs, text="Arquivo 2: Planilha Secundária").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame_inputs, textvariable=self.entrada2_var, width=60).grid(row=1, column=1, pady=5, padx=5)
        ttk.Button(frame_inputs, text="Selecionar", command=lambda: self.selecionar_arquivos("secundaria")).grid(row=1, column=2, padx=5)

        self.status_label = ttk.Label(self.root, text="")

        # Frame para Listboxes
        self.frame_listbox = ttk.LabelFrame(self.root, text="Seleção de Nomes Vulgares", padding=(10, 10))

        ttk.Label(self.frame_listbox, text="Pesquisar:").grid(row=0, column=0, sticky="w")
        pesquisa_entry = ttk.Entry(self.frame_listbox, textvariable=self.pesquisa_var, width=40)
        pesquisa_entry.grid(row=0, column=1, padx=10, pady=5)
        pesquisa_entry.bind("<KeyRelease>", self.pesquisar_nomes)

        self.listbox_nomes_vulgares = tk.Listbox(self.frame_listbox, selectmode=tk.SINGLE, width=40, height=20)
        self.listbox_nomes_vulgares.bind("<<ListboxSelect>>", self.adicionar_selecao)
        self.listbox_nomes_vulgares.grid(row=1, column=0, padx=10, pady=10)

        self.listbox_selecionados = tk.Listbox(self.frame_listbox, width=40, height=20)
        self.listbox_selecionados.grid(row=1, column=1, padx=10, pady=10)

        ttk.Button(self.frame_listbox, text="Remover Último", command=self.remover_ultimo_selecionado).grid(row=2, column=0, pady=10)
        ttk.Button(self.frame_listbox, text="Limpar Lista", command=self.limpar_lista_selecionados).grid(row=2, column=1, pady=10)

        # Frame para o botão de processamento
        self.frame_secundario = ttk.Frame(self.root, padding=(10, 10))
        ttk.Button(self.frame_secundario, text="Processar Planilhas", command=self.processar_planilhas, width=40).pack(pady=10)

        self.frame_listbox.pack_forget()  # Inicialmente escondido
        self.frame_secundario.pack_forget()  # Inicialmente escondido

    def selecionar_arquivos(self, tipo):
        """Seleciona os arquivos das planilhas."""
        global planilha_principal, planilha_secundaria

        arquivo = filedialog.askopenfilename(
            title=f"Selecione a planilha {'principal' if tipo == 'principal' else 'secundária'}",
            filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
        )
        if arquivo:
            if tipo == "principal":
                self.entrada1_var.set(arquivo)
                threading.Thread(target=self.carregar_planilha_principal, args=(arquivo,)).start()
            elif tipo == "secundaria":
                self.entrada2_var.set(arquivo)
                threading.Thread(target=self.carregar_planilha_secundaria, args=(arquivo,)).start()

    def carregar_planilha_principal(self, arquivo):
        """Carrega a planilha principal e exibe os nomes vulgares."""
        global planilha_principal, nomes_vulgares
        try:
            self.status_label.config(text="Carregando inventário principal...")
            self.status_label.pack(pady=10)

            planilha_principal = pd.read_excel(arquivo, engine="openpyxl")
            colunas_existentes = [col for col in planilha_principal.columns if col in ["Nome Vulgar"]]
            if not colunas_existentes:
                raise ValueError("A planilha principal não possui a coluna 'Nome Vulgar'.")
            nomes_vulgares = sorted(planilha_principal["Nome Vulgar"].dropna().unique())
            self.atualizar_listbox_nomes("")  # Inicializa a Listbox com todos os nomes

            self.frame_listbox.pack(pady=10)
            self.frame_secundario.pack(pady=10)
        finally:
            self.status_label.pack_forget()

    def carregar_planilha_secundaria(self, arquivo):
        """Carrega a planilha secundária."""
        global planilha_secundaria
        planilha_secundaria = pd.read_excel(arquivo, engine="openpyxl")

    def atualizar_listbox_nomes(self, filtro):
        """Atualiza a Listbox com nomes vulgares que atendem ao filtro."""
        self.listbox_nomes_vulgares.delete(0, tk.END)
        for nome in nomes_vulgares:
            if filtro.lower() in nome.lower():
                self.listbox_nomes_vulgares.insert(tk.END, nome)

    def adicionar_selecao(self, event):
        """Adiciona um nome à lista de selecionados ao clicar."""
        global nomes_selecionados
        selecao = self.listbox_nomes_vulgares.curselection()
        if selecao:
            nome = self.listbox_nomes_vulgares.get(selecao[0])
            if nome not in nomes_selecionados:
                nomes_selecionados.append(nome)
                self.atualizar_listbox_selecionados()

    def atualizar_listbox_selecionados(self):
        """Atualiza a Listbox com os nomes selecionados."""
        self.listbox_selecionados.delete(0, tk.END)
        for nome in nomes_selecionados:
            self.listbox_selecionados.insert(tk.END, nome)

    def remover_ultimo_selecionado(self):
        """Remove o último nome adicionado à lista de selecionados."""
        if nomes_selecionados:
            nomes_selecionados.pop()
            self.atualizar_listbox_selecionados()

    def limpar_lista_selecionados(self):
        """Limpa todos os nomes da lista de selecionados."""
        nomes_selecionados.clear()
        self.atualizar_listbox_selecionados()

    def processar_planilhas(self):
        """Processa os dados das planilhas."""
        global planilha_principal, planilha_secundaria
        if planilha_principal is None:
            tk.messagebox.showerror("Erro", "Por favor, carregue a planilha principal.")
            return
        if planilha_secundaria is None:
            tk.messagebox.showerror("Erro", "Por favor, carregue a planilha secundária.")
            return
        print(f"Processando planilhas com os nomes selecionados: {nomes_selecionados}")



