import tkinter as tk
import configparser

# Variáveis globais para os valores (inicializadas como None)
valor1 = None
valor2 = None
valor3 = None
valor4 = None

# Função para carregar os valores do arquivo de configuração
def carregar_valores():
    global valor1, valor2, valor3, valor4  # Definindo que vamos usar as variáveis globais
    config = configparser.ConfigParser()
    try:
        config.read('config.ini')
        # Atribuindo os valores lidos ao invés de declarar locais
        valor1 = config.get('DEFAULT', 'dapmax', fallback="0.5")
        valor2 = config.get('DEFAULT', 'dapmin', fallback="2.0")
        valor3 = config.get('DEFAULT', 'qf', fallback="3")
        valor4 = config.get('DEFAULT', 'alt', fallback="0")

        # print(f"Valor 1: {valor1}")
        # print(f"Valor 2: {valor2}")
        # print(f"Valor 3: {valor3}")
        # print(f"Valor 4: {valor4}")
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
    """Abre uma nova janela com inputs de valores padrão e bloqueia a janela principal."""    
    global valor1, valor2, valor3, valor4
    # Carregar os valores do arquivo de configuração
    carregar_valores()  # Agora as variáveis globais valor1, valor2, valor3 e valor4 serão preenchidas corretamente

    # Função para confirmar os valores inseridos
    def confirmar_valores():
        """Função executada ao clicar no botão 'Confirmar'."""
        global valor1, valor2, valor3, valor4

        # Atualiza as variáveis globais com os novos valores
        valor1 = input_valor1.get()
        valor2 = input_valor2.get()
        valor3 = input_valor3.get()
        valor4 = input_valor4.get()

        # Exibe os valores no terminal
        
        
        # Salva os valores no arquivo de configuração
        salvar_valores(valor1, valor2, valor3, valor4)

        # Fecha a janela ao confirmar
        janela_padrao.destroy()

    # Criando a janela de valores padrões
    janela_padrao = tk.Toplevel(root)
    janela_padrao.title("Critérios para REM")
    janela_padrao.geometry("400x500")
    janela_padrao.resizable(False, False)  # Impede o redimensionamento da janela

    # Bloqueia a janela principal enquanto a janela secundária estiver aberta
    janela_padrao.transient(root)  # A janela secundária estará "associada" à janela principal
    janela_padrao.grab_set()  # Bloqueia a interação com a janela principal

    # Adiciona campos de entrada de dados
    tk.Label(janela_padrao, text="DAP < :", font=("Arial", 12)).pack(pady=5)
    input_valor1 = tk.Entry(janela_padrao, font=("Arial", 12))
    input_valor1.insert(0, valor1)  # Carrega o valor salvo no arquivo
    input_valor1.pack(pady=5)

    tk.Label(janela_padrao, text="DAP >= :", font=("Arial", 12)).pack(pady=5)
    input_valor2 = tk.Entry(janela_padrao, font=("Arial", 12))
    input_valor2.insert(0, valor2)  # Carrega o valor salvo no arquivo
    input_valor2.pack(pady=5)

    tk.Label(janela_padrao, text="QF :", font=("Arial", 12)).pack(pady=5)
    input_valor3 = tk.Entry(janela_padrao, font=("Arial", 12))
    input_valor3.insert(0, valor3)  # Carrega o valor salvo no arquivo
    input_valor3.pack(pady=5)

    tk.Label(janela_padrao, text="ALT  > :", font=("Arial", 12)).pack(pady=5)
    input_valor4 = tk.Entry(janela_padrao, font=("Arial", 12))
    input_valor4.insert(0, valor4)  # Carrega o valor salvo no arquivo
    input_valor4.pack(pady=5)

    # Botão de confirmação
    tk.Button(
        janela_padrao,
        text="Confirmar",
        font=("Arial", 12),
        bg="lightblue",
        command=confirmar_valores,
        width=12,
        height=2,
    ).pack(pady=10)
