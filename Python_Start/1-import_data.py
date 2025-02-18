import pandas as pd

# 1- Importando Dados
data = pd.read_excel("/VendaCarros.xlsx")

print(data)

# 2- Lista os primeiros registros
print(data.head())

# 3- Lista os últimos registros
print(data.tail())

# 4- Contagem de valores por fabricantes
print(data["Fabricante"].value_counts())