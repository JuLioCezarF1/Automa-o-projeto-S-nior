import padronizadorPlanilha
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook


root = tk.Tk()
root.withdraw()
caminho_planilha = filedialog.askopenfilename(title="Selecione a planilha para padronizar")

dataFrame = pd.read_excel(caminho_planilha, header=2, sheet_name='Concluídos(Aprovados)')
df = pd.DataFrame(dataFrame)

padronizadorPlanilha.novasColunas(df)
padronizadorPlanilha.reembolsavel(df)
padronizadorPlanilha.mapCC(df)
padronizadorPlanilha.mapCorrespondente(df)

df.to_excel("PlanilhaPadronizada.xlsx")
print(df.dtypes)
padronizadorPlanilha.formatacaoDatas("PlanilhaPadronizada.xlsx")




