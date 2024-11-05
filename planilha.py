import os

os.system("pip install openpyxl")

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


f = []
for i in os.listdir():
	if i.endswith(".xlsx" or ".xls" or ".csv"):
		f.append(i)
		
print("Planilhas: ")
for o, b in enumerate(f, start=1):
		print(f'{o}. {b}')

print("")

planilha = int(input("Digite o número da planilha: ")) - 1

print("")

workbook = load_workbook(f[planilha])
sheet = workbook.active 

max_row = sheet.max_row
max_column = sheet.max_column

k = []
for i in range(1, max_column+1):
	k.append(get_column_letter(i))
	
print("Colunas: ")
for j, l in enumerate(k, start=1):
	print(f"{j}. {l}")

nome_column = int(input("\nDigite o número da coluna em que está o nome: ")) - 1
cpf_column = int(input("Digite o número da coluna em que está o CPF: ")) - 1

print("")

s = sheet.cell(1, cpf_column + 1).value
if "str" in str(type(s)):
	sheet.delete_rows(1)
	
for a in sheet.values:
	print(a[nome_column], a[cpf_column])
