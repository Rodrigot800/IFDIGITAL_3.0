import tkinter as tk

# Criar a janela principal
app = tk.Tk()
app.title("Minha Janela")  # Define o título da janela
app.geometry("400x300")  # Define tamanho da janela (largura x altura)
label1 = tk.Label(app, text="Linha 0, Coluna 0")
label1.grid(row=0, column=0)

label2 = tk.Label(app, text="Linha 0, Coluna 1")
label2.grid(row=5, column=1)

label3 = tk.Label(app, text="linha 3, coluna 3")
label3.grid(row=5, column=3)
# Executar a aplicação
app.mainloop()
