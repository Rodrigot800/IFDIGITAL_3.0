import tkinter as tk

def abrir_splash():
    splash = tk.Tk()
    splash.title("Carregando...")
    splash.geometry("300x150")
    
    label = tk.Label(splash, text="Carregando... Aguarde", font=("Arial", 12))
    label.pack(expand=True)
    
    splash.update()
    return splash