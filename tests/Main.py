import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import os

# Criar janela principal
app = tk.Tk()
app.title("IFDIGITAL 3.0")

# Dimensões da janela
largura_janela = 1600
altura_janela = 900

# Obter largura e altura da tela
largura_tela = app.winfo_screenwidth()
altura_tela = app.winfo_screenheight()

# Calcular coordenadas para centralizar a janela
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

# Definir geometria centralizada
app.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

# Definir estilo global para os widgets ttk
style = ttk.Style()
style.configure('.', font=("Arial", 10))

# Caminho correto da imagem
caminho_imagem = os.path.join("src", "01florest.png")

if not os.path.exists(caminho_imagem):
    print(f"Erro: Arquivo {caminho_imagem} não encontrado!")

# Variável global para manter a imagem na memória
global fundo_tk  
imagem_fundo = Image.open(caminho_imagem)
imagem_fundo = imagem_fundo.resize((largura_janela, altura_janela), Image.Resampling.LANCZOS)
fundo_tk = ImageTk.PhotoImage(imagem_fundo)

# Criar um canvas e definir a imagem de fundo
canvas = tk.Canvas(app, width=largura_janela, height=altura_janela)
canvas.pack(fill="both", expand=True)
canvas.create_image(0, 0, image=fundo_tk, anchor="nw")  # Define a imagem de fundo

# **Enviar a imagem para o fundo do Canvas**
canvas.lower("all")

# Criar um frame onde os elementos ficarão por cima da imagem
frame_conteudo = tk.Frame(app, bg="white")
frame_conteudo.place(relx=0.5, rely=0.5, anchor="center", width=400, height=200)

# Adicionar um botão de teste dentro do frame
btn_teste = ttk.Button(frame_conteudo, text="Botão Teste")
btn_teste.pack(pady=20)

# Executar a aplicação
app.mainloop()
