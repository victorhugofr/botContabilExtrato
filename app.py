import openpyxl
from datetime import datetime
import pandas as pd

from time import sleep
import tkinter as tk
from tkinter import filedialog

def selecionar_arquivo():
    global local_arquivo
    arquivo = filedialog.askopenfilename(initialdir="/", title="Selecione o arquivo", filetypes=(("Arquivos de texto", "*.xls"), ("Todos os arquivos", "*.*")))
    if arquivo:
        local_arquivo=arquivo
        label_local_arquivo.config(text="Local do arquivo selecionado: " + arquivo)
        janela.destroy()
    else:
        label_local_arquivo.config(text="Nenhum arquivo selecionado")

# Criar janela
janela = tk.Tk()
janela.title("Selecionar Arquivo")

# Criar botão para selecionar arquivo
botao_selecionar = tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo)
botao_selecionar.pack(pady=10)

# Label para exibir o local do arquivo selecionado
label_local_arquivo = tk.Label(janela, text="")
label_local_arquivo.pack()

# Rodar aplicação
janela.mainloop()
#workbook = openpyxl.load_workbook(local_arquivo)
workbook = pd.read_excel(local_arquivo)
novoWb = openpyxl.Workbook()
ws = novoWb.active
totalDebitos=0
totalCreditos=0
data_hora_atual = datetime.now()
tipo=''
data_hora_formatada = data_hora_atual.strftime("%Y%m%d %H%M%S")
for linha,row in workbook.iterrows():
    if(pd.notna(workbook.iloc[linha,0]) and 'Movimento' in workbook.iloc[linha,0]):
        tipo = workbook.iloc[linha,4]
    if(pd.notna(workbook.iloc[linha,0]) and 'Total de débitos' in workbook.iloc[linha,0] and totalDebitos==0):
        linhaAInserir=["Total de débitos",workbook.iloc[linha,18]]
        totalDebitos=1
        ws.append(linhaAInserir)
    if(pd.notna(workbook.iloc[linha,0]) and 'Total de créditos' in workbook.iloc[linha,0] and totalCreditos==0):
        linhaAInserir=["Total de créditos",workbook.iloc[linha,18]]
        totalCreditos=1
        ws.append(linhaAInserir)
    if(pd.notna(workbook.iloc[linha,1]) and 'CFOP' in workbook.iloc[linha,1]):
        linhaAInserir=[workbook.iloc[linha,1],workbook.iloc[linha,14],tipo]
        ws.append(linhaAInserir)
       
novoWb.save(data_hora_formatada+".xlsx")