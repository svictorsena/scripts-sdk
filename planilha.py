from customtkinter import *
import os
from tkinter import *
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

import keyboard
import pyautogui
import time

set_appearance_mode("dark") 
set_default_color_theme("blue")

t = CTk()
t.geometry("800x500")
t.resizable(False, False)

entrada = CTkLabel(t, width=200, text="Selecione a planilha", font=("Arial", 20, "bold"))
entrada.pack(pady=20)

f = [i for i in os.listdir() if i.endswith(".xlsx")]

def io():
	global cpf_column, nome_column, sheet, button
	entrada.configure(text="bagulho aqui")
	s = sheet.cell(1, cpf_column + 1).value
	if "str" in str(type(s)):
		sheet.delete_rows(1)
	
	time.sleep(10)
	for i in sheet.values:
		pyautogui.write(str(i[cpf_column]))
		time.sleep(1)
		pyautogui.press("tab")
		time.sleep(1)
		pyautogui.write(str(i[nome_column]))
		time.sleep(1)
		pyautogui.press("tab")
		pyautogui.press("tab")
		if keyboard.is_pressed("esc"):
			break

def oii(event):
	global cpf_column, nome_column, button, cb2
	nome_column = column_index_from_string(event) - 1
	
	cb2.pack_forget()
	entrada.configure(text="Tudo certinho! Só apertar o botão abaixo para iniciar o processo.")
	
	button = CTkButton(t, text="iniciar", width=100, command=io).pack()

def oi(column):
    global cb1, k, cpf_column, cb2
    cpf_column = column_index_from_string(column) - 1
    entrada.configure(text="Selecione a coluna em que está o nome")
    cb1.pack_forget()
    
    k = [x for x in k if x != column]
	
    cb2 = CTkComboBox(t, width=300, values=k, command=oii)
    cb2.set("--selecione--")
    cb2.pack()

def value(event):
    global k, cb1, sheet# Tornando k uma variável global
    planilha = cb.get()
    wb = load_workbook(planilha)
    sheet = wb.active
    
    max_column = sheet.max_column
    
    k = [get_column_letter(i) for i in range(1, max_column + 1)]
    
    entrada.configure(text="Selecione a coluna em que está o CPF")
    cb.pack_forget()

    cb1 = CTkComboBox(t, width=300, values=k, command=oi)
    cb1.set("--selecione--")
    cb1.pack()

cb = CTkComboBox(t, width=300, values=f, command=value)
cb.set("--selecione--")
cb.pack()

t.mainloop()
