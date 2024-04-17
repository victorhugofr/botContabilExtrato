import openpyxl
from datetime import datetime

from time import sleep
import tkinter as tk
from tkinter import filedialog

def selecionar_arquivo():
    global local_arquivo
    arquivo = filedialog.askopenfilename(initialdir="/", title="Selecione o arquivo", filetypes=(("Arquivos de texto", "*.xlsx"), ("Todos os arquivos", "*.*")))
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
workbook = openpyxl.load_workbook(local_arquivo)
sheet_produtos=workbook['dem']
novoWb = openpyxl.Workbook()
ws = novoWb.active
totalDebitos=0
totalCreditos=0
data_hora_atual = datetime.now()
data_hora_formatada = data_hora_atual.strftime("%Y%m%d %H%M%S")
for linha in sheet_produtos.iter_rows(min_row=9):
    if(linha[0].value is not None and 'Total de débitos' in linha[0].value and totalDebitos==0):
        linhaAInserir=["Total de débitos",linha[18].value]
        totalDebitos=1
        ws.append(linhaAInserir)
    if(linha[0].value is not None and 'Total de créditos' in linha[0].value and totalCreditos==0):
        linhaAInserir=["Total de créditos",linha[18].value]
        totalCreditos=1
        ws.append(linhaAInserir)
    if(linha[1].value is not None and 'CFOP' in linha[1].value):
        linhaAInserir=[linha[1].value,linha[14].value]
        ws.append(linhaAInserir)
       
novoWb.save(data_hora_formatada+".xlsx")