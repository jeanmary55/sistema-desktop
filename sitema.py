import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import requests
from tkinter.filedialog import askopenfilename
import numpy 
import pandas as pd
from datetime import datetime


link = 'https://economia.awesomeapi.com.br/json/all'
requisicao = requests.get(link)
dicionario_moedas = requisicao.json()


def pegarCotacao():
    moeda = combobox_moeda.get()
    data = calendar_moeda.get()
    ano = data[-4:]
    mes = data[3:5]
    dia = data[0:2]
    
   
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    
    requisicao = requests.get(link)
    dicionario = requisicao.json()
    cotacao = dicionario[0]['bid']
    
    
    cotacao_no_dia['text'] = f'A cotação da {moeda} no dia {data} é de {cotacao}'


    
    
    

def selecaoArquivo() :
    arquivo = askopenfilename(title='Selecionar um arquivo moeda')
    var_button_seleciona.set(arquivo)

    if arquivo :
        selecao_arquivo['text'] = f'o arquivo selecionado é {arquivo}'




def atualisar() :
    df = pd.read_excel(var_button_seleciona.get())
    moedas = df.iloc[:, 0]
    data_init = calendar_data_inicial.get()
    data_fina = calendar_data_final.get()
   

    ano_initial = data_init[-4:]
    mes_intial = data_init[3:5]
    dia_initial = data_init[0:2]

    ano_final = data_fina[-4:]
    mes_final = data_fina[3:5]
    dia_final = data_fina[0:2]
   
    for moeda in moedas :
        try:
            link = link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?" \
                    f"start_date={ano_initial}{mes_intial}{dia_initial}&" \
                    f"end_date={ano_final}{mes_final}{dia_final}"


            requisicao = requests.get(link)
            cotacoes = requisicao.json()
            
            for cotacao in cotacoes :
                bid = float(cotacao['bid'])
                timstap = int(cotacao['timestamp'])
                data = datetime.fromtimestamp(timstap)
                data = data.strftime('%d/%m/%Y')

                if data not in df:
                    df[data] = numpy.nan
                
                df.loc[df.iloc[:, 0]== moeda, data] = bid
                df.to_excel('texte2.xlsx')

                mensagem['text'] = 'atualisar com sucesso'
        except :
            mensagem['text'] = 'arquivo errado'



    



lista_moeda = list(dicionario_moedas.keys())
janela = tk.Tk()

janela.columnconfigure([0,2], weight=1)

cotacao_moeda = tk.Label(text= 'Cotação de 1 moeda específica',borderwidth=2, relief='solid',padx=10, pady=10)
cotacao_moeda.grid(row=0, column=0,padx=10, pady=10, columnspan=3, sticky='nwes')

selecao_moeda = tk.Label(text= 'Selecione a moeda que desejar consultar:',padx= 10, pady=10 )
selecao_moeda.grid(row=1, column=0, columnspan=2, sticky='nwes')

combobox_moeda = ttk.Combobox(values=lista_moeda)
combobox_moeda.grid(row= 1, column=2,padx=10, pady=10, sticky='nsew')

dia_cotacao = tk.Label(text='Selecionar  o dia  que  você quer pegar cotação ', anchor='e')
dia_cotacao.grid(row= 2, column=0,columnspan= 2,padx=10, pady=10, sticky='nsew')

calendar_moeda = DateEntry(year = 2021, locale='pt_br')
calendar_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nsew')

cotacao_no_dia = tk.Label(text='A cotação da moeda tal no dia tal é de tanto')
cotacao_no_dia.grid(row=3, column= 0, columnspan= 2,padx=10, pady=10, sticky='nsew')

botao_pegar_cotacao = tk.Button(text= 'Pegar Cotação', command=pegarCotacao)
botao_pegar_cotacao.grid(row=3, column= 2,padx=10, pady=10, sticky='nsew')

# segunda parte do  meu código
cotacoes_moeda = tk.Label(text= 'Cotação de 1 moeda específica',borderwidth=2, relief='solid',padx=10, pady=10)
cotacoes_moeda.grid(row=4, column=0, columnspan=3, sticky='nwes', padx=10, pady=10)

arquivo = tk.Label(text= 'Selecione um arquivo em Excel com as Moedas na Coluna A:')
arquivo.grid(row=5, column=0, columnspan=2, sticky='nwes', padx=10, pady=1)

var_button_seleciona = tk.StringVar()
selecao_arquivo = tk.Button(text='Clique aqui para selecionar',command= selecaoArquivo )
selecao_arquivo.grid(row=5, column=2 , sticky='nwes', padx=10, pady=1)

arquivo_mensagem = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
arquivo_mensagem.grid(row=6, column=0, columnspan=3, sticky='nwes', padx=10, pady=1)

data_initial = tk.Label(text='Data initial')
data_initial.grid(row=7, column=0, sticky='nwes', padx=10, pady=1)

data_final = tk.Label(text='Data initial')
data_final.grid(row=8, column=0, sticky='nwes', padx=10, pady=1)



calendar_data_inicial = DateEntry(year = 2021, locale='pt_br')
calendar_data_inicial.grid(row=7, column=1, sticky='nwes', padx=10, pady=1)

calendar_data_final = DateEntry(year = 2021, locale='pt_br')
calendar_data_final.grid(row=8, column=1, sticky='nwes', padx=10, pady=1)

botao_atualisar = tk.Button(text='Atualizar Cotações', command=atualisar)
botao_atualisar.grid(row=9, column=0, sticky='nwes', padx=10, pady=1) 

mensagem = tk.Label(text='Arquivo de Moedas atualizado com sucesso.')
mensagem.grid(row=9, column=1, columnspan=2,sticky='nwes', padx=10, pady=1)
janela.mainloop()