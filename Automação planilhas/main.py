import padronizadorPlanilha
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

#Nessa página está todas as chamadas das funções, optei por separar os documentos para deixar o arquivo principal(main.py) mais limpo.

root = tk.Tk()
root.withdraw()
caminho_planilha = filedialog.askopenfilename(title="Selecione a planilha para padronizar")

dataFrame = pd.read_excel(caminho_planilha, header=2, sheet_name='Concluídos(Aprovados)')
df = pd.DataFrame(dataFrame)

padronizadorPlanilha.novasColunas(df) #Essa função irá adicionar 6 novas colunas das quais estão nomeadas com o padrão da empresa, mas sinta-se a vontade para mudar o nome para o seu padrão.
padronizadorPlanilha.reembolsavel(df) #Essa função utiliza um metodo do Pandas chamado .loc que é util para verificar uma célula já existente, e preencher uma celula de acordo com condição fornecida.
padronizadorPlanilha.mapCC(df) #Essa função utiliza uma função do Pandas chamada .map() que ajuda a localizar e preencher informações em uma celula de acordo com um dicionário.
padronizadorPlanilha.mapCorrespondente(df) #Essa função também utiliza a função do Pandas .map() que irá preencher uma nova coluna criada atráves do padronizadorPlanilha.novasColunas()

df.to_excel("PlanilhaPadronizada.xlsx") #Aqui será realizado o download da planilha já com as alterações feitas, porém ainda não é a planilha final, pois a biblioteca pandas imprime as datas no formato de EUA.
print(df.dtypes)
padronizadorPlanilha.formatacaoDatas("PlanilhaPadronizada.xlsx") #O Download da planilha final terá o nome de PlanilhaPadronizada.xlsx, onde todas as formatações estarão completas e com datas no formato brasileiro.




