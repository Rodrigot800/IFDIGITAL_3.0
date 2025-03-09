import tkinter as tk
from tkinter import ttk
import pandas as pd

# Exemplo de DataFrame
data = {
    'UT': ['A', 'A', 'B', 'B'],
    'QF': [3, 1, 4, 2],
    'Volume_m3': [100, 200, 150, 250]
}
df_filtrado = pd.DataFrame(data)

def ordenar_dataframe():
    # Obtém a opção selecionada no combobox
    opcao = ordem_combobox.get()
    
    if opcao == "QF e depois Volume_m3":
        # Ordena por UT (asc), QF (desc) e Volume_m3 (asc)
        df_filtrado.sort_values(by=["UT", "QF", "Volume_m3"],
                                ascending=[True, False, True],
                                inplace=True)
    elif opcao == "Volume_m3 e depois QF":
        # Ordena por UT (asc), Volume_m3 (asc) e QF (desc)
        df_filtrado.sort_values(by=["UT", "Volume_m3", "QF"],
                                ascending=[True, True, False],
                                inplace=True)
    elif opcao == "Apenas QF":
        # Ordena por UT (asc) e QF (desc)
        df_filtrado.sort_values(by=["UT", "QF"],
                                ascending=[True, False],
                                inplace=True)
    elif opcao == "Apenas Volume_m3":
        # Ordena por UT (asc) e Volume_m3 (asc)
        df_filtrado.sort_values(by=["UT", "Volume_m3"],
                                ascending=[True, True],
                                inplace=True)
    
    # Aqui você pode atualizar a interface ou exibir o resultado
    print("DataFrame ordenado:")
    print(df_filtrado)

# Criação da interface principal
app = tk.Tk()
app.title("Controle de Ordenação")
app.geometry("400x200")

# Combobox para seleção da ordem de classificação
ordem_combobox = ttk.Combobox(app, 
                              values=["QF e depois Volume_m3", 
                                      "Volume_m3 e depois QF", 
                                      "Apenas QF", 
                                      "Apenas Volume_m3"],
                              state="readonly", width=30)
ordem_combobox.current(0)  # Define a opção padrão
ordem_combobox.pack(pady=20)

# Botão para acionar a ordenação
btn_ordenar = ttk.Button(app, text="Ordenar", command=ordenar_dataframe)
btn_ordenar.pack(pady=10)

app.mainloop()