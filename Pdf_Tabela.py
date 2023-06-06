import os
import tkinter as tk
import tkinter.filedialog as fdlg
import tkinter.messagebox as tkMessageBox
from tkinter import *
from datetime import date
import pandas as pd
from tkinter import ttk
from time import sleep
import tabula

#root = tk.Tk()
tinicial = tk.Tk()
tinicial.geometry("800x500+200+100")
tinicial.title("Extrair Tabelas de PDF - SIS")
tinicial.resizable(width=False, height=False)
tinicial['bg'] = '#49A'
tinicial.iconphoto(True, PhotoImage(file='./arquivos/quebra.png'))
image=PhotoImage(file='./arquivos/junta.png')

#importante para progressbar
s = ttk.Style() 
s.theme_use('default') 
s.configure("SKyBlue1.Horizontal.TProgressbar", foreground='DarkSeaGreen3', background='white')

data_atual = date.today()
data_atual = data_atual.strftime('%d-%m-%Y')

def juntacsv():
	progress1.start(10)
	try:
		pagina_selct_ini = qtdPaginaini.get()
		pagina_selct_fim = qtdPaginafim.get()
		lista_paginas = list(range(int(pagina_selct_ini),int(pagina_selct_fim)+1))

		#aqui seleciona os arquivos
		path = fdlg.askopenfilenames()
		df = pd.DataFrame()

		
		for f in path:
			for pg in lista_paginas:
				lista_de_tabelas = tabula.read_pdf(f,pages=pg)
				for tabela in lista_de_tabelas:
					df = df.append(tabela)


		tkMessageBox.showinfo("Selecionar Pasta", message= "Selecione Pasta para salvar!")

		#aqui seleciona a pasta a ser colocada o novo arquivo
		opcoes = {}                # as opções são definidas em um dicionário
		opcoes['initialdir'] = ''    # será o diretório atual
		opcoes['parent'] = tinicial
		opcoes['title'] = 'Diálogo que retorna o nome do diretório selecionado'
		caminhoinicial = fdlg.askdirectory(**opcoes)

		df.to_excel(caminhoinicial +'/Relatorio-RO-Tabelas-de-Planos' + data_atual + '.xlsx', index=False)

		tkMessageBox.showinfo("Junção Finalizada", message= "Realizado com sucesso!")
		progress1.stop()
	except:
		tkMessageBox.showinfo("Erro", message= "Ocorreu algum erro!")
		progress1.stop()






robozinho = Label(tinicial, image = image,width=800, height=450,bg ="white")
robozinho.grid(row=10,columnspan =10)

cmdCadastrar=Button(tinicial,bd=4,bg = 'SKyBlue1',fg='black',text='Selecionar os Arquivos',font=('arial',18,'bold'),width=25,height=2,
	command = juntacsv).grid(row=10,columnspan=10)

LblqtdPaginaIni = Label(tinicial,bd=4,bg = 'SKyBlue1',fg='black',text='Página Inicial',font=('arial',18,'bold'),width=13).grid(row=11, column=1)
qtdPaginaini = Entry(tinicial,bd=4,bg = 'SKyBlue1',fg='black',font=('arial',18,'bold'),width=4)
qtdPaginaini.grid(row=11, column=2)
qtdPaginaini.insert(0,9)

LblqtdPaginafim = Label(tinicial,bd=4,bg = 'SKyBlue1',fg='black',text='Página Final',font=('arial',18,'bold'),width=13).grid(row=11, column=3)
qtdPaginafim = Entry(tinicial,bd=4,bg = 'SKyBlue1',fg='black',font=('arial',18,'bold'),width=4)
qtdPaginafim.grid(row=11, column=4)
qtdPaginafim.insert(0,9)

progress1 =ttk.Progressbar(robozinho, orient=VERTICAL, length=450, style="SKyBlue1.Horizontal.TProgressbar",mode='determinate')
progress1.place(relx=0.005, rely = 0)


tinicial.mainloop()
