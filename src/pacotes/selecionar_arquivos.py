from tkinter import filedialog


def selecionar_arquivos():
    """Seleciona os arquivos das duas planilhas."""
    arquivo1 = filedialog.askopenfilename(
        title="Selecione a planilha de entrada (dados principais)",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if arquivo1:
        entrada1_var.set(arquivo1)

    arquivo2 = filedialog.askopenfilename(
        title="Selecione a segunda planilha (Nomes Vulgares e Científicos)",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if arquivo2:
        entrada2_var.set(arquivo2)