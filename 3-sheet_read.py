from openpyxl import load_workbook

# 1- lê pasta de trabalho e planilha
wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatorio"] #indica qual planilha quero utilizar, no caso Relatorio 

# 2- acessando um valor específico
print(sheet["A3"].value) #pegar o valor de A3 na planilha

# 3- interando valores por meio de loop
for i in range(2, 6): 
    ano = sheet["A%s" %i].value
    am = sheet["B%s" %i].value
    bt = sheet["C%s" %i].value
    print("{0} o Aston Martin vendeu {1} e o Bentley vendeu {2}".format(ano, am, bt))