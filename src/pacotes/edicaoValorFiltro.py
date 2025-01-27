import tkinter as tk

def abrir_janela_valores_padroes(root):
    """Abre uma nova janela com inputs de valores padrão."""
    
    def confirmar_valores():
        """Função executada ao clicar no botão 'Confirmar'."""
        valor1 = input_valor1.get()
        valor2 = input_valor2.get()
        valor3 = input_valor3.get()
        valor4 = input_valor4.get()

        # Exibe os valores no terminal
        print(f"Valor 1: {valor1}")
        print(f"Valor 2: {valor2}")
        print(f"Valor 3: {valor3}")
        print(f"Valor 4: {valor4}")

        # Fecha a janela ao confirmar
        janela_padrao.destroy()

    janela_padrao = tk.Toplevel(root)
    janela_padrao.title("Valores Padrões")
    janela_padrao.geometry("400x500")  # Tamanho ajustado para caber os elementos
    janela_padrao.resizable(False, False)  # Impede redimensionamento

    # Labels e inputs com valores predefinidos
    tk.Label(janela_padrao, text="Valor 1:", font=("Arial", 12)).pack(pady=5)
    input_valor1 = tk.Entry(janela_padrao, font=("Arial", 12))
    input_valor1.insert(0, "0.5")  # Valor padrão
    input_valor1.pack(pady=5)

    tk.Label(janela_padrao, text="Valor 2:", font=("Arial", 12)).pack(pady=5)
    input_valor2 = tk.Entry(janela_padrao, font=("Arial", 12))
    input_valor2.insert(0, "2.0")  # Valor padrão
    input_valor2.pack(pady=5)

    tk.Label(janela_padrao, text="Valor 3:", font=("Arial", 12)).pack(pady=5)
    input_valor3 = tk.Entry(janela_padrao, font=("Arial", 12))
    input_valor3.insert(0, "3")  # Valor padrão
    input_valor3.pack(pady=5)

    tk.Label(janela_padrao, text="Valor 4:", font=("Arial", 12)).pack(pady=5)
    input_valor4 = tk.Entry(janela_padrao, font=("Arial", 12))
    input_valor4.insert(0, "0")  # Valor padrão
    input_valor4.pack(pady=5)

    # Botão de confirmar
    tk.Button(
        janela_padrao,
        text="Confirmar",
        font=("Arial", 12),
        bg="lightblue",
        command=confirmar_valores,
        width=12,
        height=2,
    ).pack(pady=10)
