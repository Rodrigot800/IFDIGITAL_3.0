import tkinter as tk
from tkinter import filedialog, ttk

class InterfaceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Processador de Inventário e Mesclagem")
        self.root.geometry("900x700")

        # Variáveis
        self.entrada1_var = tk.StringVar()
        self.entrada2_var = tk.StringVar()
        self.pesquisa_var = tk.StringVar()
        self.nomes_selecionados = []
        self.nomes_vulgares = []  # Inicializa a lista de nomes vulgares como vazia

        # Widgets
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

        # Frame para seleção de nomes vulgares
        self.frame_listbox = ttk.LabelFrame(self.root, text="Seleção de Nomes Vulgares", padding=(10, 10))
        self.frame_listbox.pack(pady=10, padx=10, fill="x")

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
        ttk.Button(self.root, text="Processar Planilhas", command=self.processar_planilhas, width=40).pack(pady=20)

    def selecionar_arquivos(self, tipo):
        """Seleciona os arquivos das planilhas."""
        arquivo = filedialog.askopenfilename(
            title=f"Selecione a planilha {'principal' if tipo == 'principal' else 'secundária'}",
            filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
        )
        if arquivo:
            if tipo == "principal":
                self.entrada1_var.set(arquivo)
                # Aqui você pode carregar a planilha principal
            elif tipo == "secundaria":
                self.entrada2_var.set(arquivo)
                # Aqui você pode carregar a planilha secundária

    def pesquisar_nomes(self, event=None):
        """Filtra os nomes vulgares na Listbox."""
        filtro = self.pesquisa_var.get().lower()
        self.listbox_nomes_vulgares.delete(0, tk.END)
        for nome in self.nomes_vulgares:  # 'nomes_vulgares' será preenchido após carregar a planilha
            if filtro in nome.lower():
                self.listbox_nomes_vulgares.insert(tk.END, nome)

    def adicionar_selecao(self, event):
        """Adiciona um nome à lista de selecionados."""
        selecao = self.listbox_nomes_vulgares.curselection()
        if selecao:
            nome = self.listbox_nomes_vulgares.get(selecao[0])
            if nome not in self.nomes_selecionados:
                self.nomes_selecionados.append(nome)
                self.atualizar_listbox_selecionados()

    def atualizar_listbox_selecionados(self):
        """Atualiza a Listbox com os nomes selecionados."""
        self.listbox_selecionados.delete(0, tk.END)
        for nome in self.nomes_selecionados:
            self.listbox_selecionados.insert(tk.END, nome)

    def remover_ultimo_selecionado(self):
        """Remove o último nome adicionado à lista de selecionados."""
        if self.nomes_selecionados:
            self.nomes_selecionados.pop()
            self.atualizar_listbox_selecionados()

    def limpar_lista_selecionados(self):
        """Limpa todos os nomes da lista de selecionados."""
        self.nomes_selecionados.clear()
        self.atualizar_listbox_selecionados()

    def processar_planilhas(self):
        """Inicia o processamento das planilhas."""
        print(f"Processando as planilhas com os nomes selecionados: {self.nomes_selecionados}")
        # Aqui você chamará a função de processamento com os arquivos selecionados e os nomes

    def atualizar_listbox_nomes(self, filtro):
      """Atualiza a Listbox com nomes vulgares que atendem ao filtro."""
    self.listbox_nomes_vulgares.delete(0, tk.END)  # Limpa a Listbox
    for nome in self.nomes_vulgares:
        if filtro.lower() in nome.lower():  # Aplica o filtro
            self.listbox_nomes_vulgares.insert(tk.END, nome)  # Adiciona os nomes ao Listbox


