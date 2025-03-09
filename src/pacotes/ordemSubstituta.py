import pandas as pd
import tkinter as tk
from tkinter import ttk

def ordenar_dataframe(df: pd.DataFrame, criterio: str) -> pd.DataFrame:
    """
    Ordena o DataFrame com base no critério selecionado.

    Parâmetros:
      df: DataFrame a ser ordenado.
      criterio: String que define a ordenação. Pode ser:
          - "QF e depois Volume_m3": Ordena por UT (asc), QF (desc) e Volume_m3 (asc)
          - "Volume_m3 e depois QF": Ordena por UT (asc), Volume_m3 (asc) e QF (desc)
          - "Apenas QF": Ordena por UT (asc) e QF (desc)
          - "Apenas Volume_m3": Ordena por UT (asc) e Volume_m3 (asc)

    Retorna:
      O mesmo DataFrame ordenado.
    """
    if criterio == "QF e depois Volume_m3":
        df.sort_values(by=["UT", "QF", "Volume_m3"],
                       ascending=[True, False, True],
                       inplace=True)
    elif criterio == "Volume_m3 e depois QF":
        df.sort_values(by=["UT", "Volume_m3", "QF"],
                       ascending=[True, True, False],
                       inplace=True)
    elif criterio == "Apenas QF":
        df.sort_values(by=["UT", "QF"],
                       ascending=[True, False],
                       inplace=True)
    elif criterio == "Apenas Volume_m3":
        df.sort_values(by=["UT", "Volume_m3"],
                       ascending=[True, True],
                       inplace=True)
    else:
        print("Critério não reconhecido. Nenhuma ordenação aplicada.")
    
    return df

class OrdenadorFrame(ttk.Frame):
    """
    Um frame Tkinter que contém um Combobox para seleção da ordenação
    e um botão que aplica a ordenação ao DataFrame fornecido.
    """
    def __init__(self, master, dataframe: pd.DataFrame, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.dataframe = dataframe
        self._criar_widgets()

    def _criar_widgets(self):
        # Combobox para seleção do critério de ordenação
        self.combobox = ttk.Combobox(
            self,
            values=[
                "QF e depois Volume_m3",
                "Volume_m3 e depois QF",
                "Apenas QF",
                "Apenas Volume_m3"
            ],
            state="readonly",
            width=30
        )
        self.combobox.current(0)  # Define um valor padrão
        self.combobox.pack(pady=10)

        # Botão para aplicar a ordenação
        self.btn_ordenar = ttk.Button(self, text="Ordenar", command=self.ordenar)
        self.btn_ordenar.pack(pady=10)

    def ordenar(self):
        # Recupera o critério escolhido e ordena o DataFrame
        criterio = self.combobox.get()
        ordenar_dataframe(self.dataframe, criterio)
        print("DataFrame ordenado:")
        print(self.dataframe)