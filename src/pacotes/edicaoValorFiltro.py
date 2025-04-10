import tkinter as tk
import configparser
import os
import sys

# Variáveis globais para os valores (inicializadas como None)
valor1 = None
valor2 = None
valor3 = None
valor4 = None


# Função para garantir o caminho correto tanto no executável quanto no script de desenvolvimento
def resource_path(relative_path):
    """ Garante o caminho certo tanto no executável quanto em desenvolvimento """
    try:
        base_path = sys._MEIPASS  # Caso o script esteja executando como .exe (PyInstaller)
    except Exception:
        base_path = os.path.abspath(".")  # Caso esteja executando como script

    return os.path.join(base_path, relative_path)

# Função para carregar os valores do arquivo de configuração
def carregar_valores():
    global valor1, valor2, valor3, valor4  # Usando as variáveis globais
    config = configparser.ConfigParser()
    try:
        config.read('config.ini')
        valor1 = config.get('DEFAULT', 'dapmax', fallback="0.5")
        valor2 = config.get('DEFAULT', 'dapmin', fallback="2.0")
        valor3 = config.get('DEFAULT', 'qf', fallback="3")
        valor4 = config.get('DEFAULT', 'alt', fallback="0")
    except FileNotFoundError:
        valor1 = "0.5"
        valor2 = "2.0"
        valor3 = "3"
        valor4 = "0"
    
    return valor1, valor2, valor3, valor4

# Função para salvar os valores no arquivo de configuração
def salvar_valores(valor1, valor2, valor3, valor4):
    config = configparser.ConfigParser()
    config['DEFAULT'] = {
        'dapmax': valor1,
        'dapmin': valor2,
        'qf': valor3,
        'alt': valor4
    }
    
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

def abrir_janela_valores_padroes(root):
    global valor1, valor2, valor3, valor4
    carregar_valores()  # Carrega os valores e preenche as variáveis globais

    def confirmar_valores():
        global valor1, valor2, valor3, valor4
        valor1 = input_valor1.get()
        valor2 = input_valor2.get()
        valor3 = input_valor3.get()
        if toggle_alt.get() == 1:
            valor4 = input_valor4.get()
        else:
            valor4 = "0"
        salvar_valores(valor1, valor2, valor3, valor4)
        janela_padrao.destroy()

    # Cria a janela de configurações
    janela_padrao = tk.Toplevel(root)
    janela_padrao.title("Critérios para REM")
    janela_padrao.geometry("300x350")
    janela_padrao.resizable(False, False)
    janela_padrao.transient(root)
    janela_padrao.grab_set()

    # Caminho para o ícone da janela
    icone_path = resource_path("src/img/icoGreenFlorest.ico")
    janela_padrao.iconbitmap(icone_path)

    # Frame para os inputs principais
    frame_inputs = tk.Frame(janela_padrao)
    frame_inputs.pack(pady=10, fill='both', expand=True)

    # Input para DAP <
    tk.Label(frame_inputs, text="DAP < :", font=("Arial", 11)).pack(pady=5)
    input_valor1 = tk.Entry(frame_inputs, font=("Arial", 11))
    input_valor1.insert(0, valor1)
    input_valor1.pack(pady=5)

    # Input para DAP >=
    tk.Label(frame_inputs, text="DAP >= :", font=("Arial", 11)).pack(pady=5)
    input_valor2 = tk.Entry(frame_inputs, font=("Arial", 11))
    input_valor2.insert(0, valor2)
    input_valor2.pack(pady=5)

    # Input para QF
    tk.Label(frame_inputs, text="QF >= :", font=("Arial", 11)).pack(pady=5)
    input_valor3 = tk.Entry(frame_inputs, font=("Arial", 11))
    input_valor3.insert(0, valor3)
    input_valor3.pack(pady=5)

    # Container para os controles do H (checkbutton e input)
    frame_alt_container = tk.Frame(frame_inputs)
    frame_alt_container.pack(pady=5, fill='x')

    toggle_alt = tk.IntVar()
    if valor4 != "0":
        toggle_alt.set(1)
    else:
        toggle_alt.set(0)
    
    checkbutton_alt = tk.Checkbutton(
        frame_alt_container,
        text="H > :",
        variable=toggle_alt,
        font=("Arial", 11)
    )
    checkbutton_alt.pack(pady=5)

    # Frame para o campo H
    frame_alt = tk.Frame(frame_alt_container)
    input_valor4 = tk.Entry(frame_alt, font=("Arial", 11))
    input_valor4.insert(0, valor4)
    input_valor4.pack(pady=5)

    # Função para alternar a visibilidade do campo H
    def toggle_alt_input():
        if toggle_alt.get() == 1:
            frame_alt.pack(pady=5, fill='x')
        else:
            frame_alt.pack_forget()
            input_valor4.delete(0, tk.END)
            input_valor4.insert(0, "0")
    
    # Associa a função ao checkbutton
    checkbutton_alt.config(command=toggle_alt_input)
    # Exibe o campo H se estiver ativo inicialmente
    if toggle_alt.get() == 1:
        frame_alt.pack(pady=5, fill='x')

    # Frame para o botão de confirmação (fixo no final)
    frame_confirm = tk.Frame(janela_padrao)
    frame_confirm.pack(pady=10)
    tk.Button(
        frame_confirm,
        text="Confirmar",
        font=("Arial", 11),
        command=confirmar_valores,
        width=10,
        height=2,
    ).pack()

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Janela Principal")
    tk.Button(root, text="Abrir Configurações", command=lambda: abrir_janela_valores_padroes(root)).pack(pady=20)
    root.mainloop()
