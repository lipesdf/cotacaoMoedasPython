import tkinter as tk
import requests
import pandas as pd
import numpy as np

from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
from datetime import datetime

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')

# Transformou dicionario JSON em DICIONARIO 
dicionarioMoedas = requisicao.json()

listaMoedas = list(dicionarioMoedas.keys())


def pegarCotacao():
    moeda = comboboxSelecionarMoeda.get()
    dataCotacao = calendarioMoeda.get()
    ano = dataCotacao[6:10]
    mes = dataCotacao[3:5]
    dia = dataCotacao[:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisaoMoeda = requests.get(link)
    cotacao = requisaoMoeda.json()
    valorMoeda = cotacao[0]['bid']
    labelTextoCotacao['text'] = f'A cotação da moeda {moeda} no dia {dataCotacao} foi de: R${valorMoeda}.'


def selecionarArquivo():
    caminhoArquivo = askopenfilename(title="Selecione o arquivo da Moeda")
    varCaminho.set(caminhoArquivo)
    if caminhoArquivo:
        labelArquivoSelecionado['text'] = f"Arquivo selecionado: {caminhoArquivo}"


def atualizarCotacao():
    try:
        df = pd.read_excel(varCaminho.get())

        dataInicial = calendarioDataInicial.get()
        dataFinal = calendarioDataFinal.get()

        moedas = df.iloc[:, 0]

        anoInicial = dataInicial[6:10]
        mesInicial = dataInicial[3:5]
        diaInicial = dataInicial[:2]

        anoFinal = dataFinal[6:10]
        mesFinal = dataFinal[3:5]
        diaFinal = dataFinal[:2]

        
        for moeda in moedas:
            link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={anoInicial}{mesInicial}{diaInicial}&end_date={anoFinal}{mesFinal}{diaFinal}'
            requisaoMoeda = requests.get(link)
            cotacoes = requisaoMoeda.json()
            
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')

                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:,0] == moeda , data] = bid

        df.to_excel('Moedas.xlsx')
        labelAtualizarCotacao['text'] = 'Arquivo atualizado com sucesso.'
    except:
        labelAtualizarCotacao['text'] = 'Selecione um arquivo Excel no formato correto'


janela = tk.Tk()

janela.title("Ferramenta de cotação de moedas")

labelCotacaoMoeda = tk.Label(text='Cotação de 1 moeda específica', borderwidth=2, relief='solid')
labelCotacaoMoeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

# anchor = alinhamento
labelSelecionarMoeda = tk.Label(text='Selecionar moeda', anchor='w')
labelSelecionarMoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

comboboxSelecionarMoeda = ttk.Combobox(values=listaMoedas)
comboboxSelecionarMoeda.grid(row=1, column=2, padx=10, pady=10, sticky='nswe')

labelSelecionarDia = tk.Label(text='Selecione o dia que deseja pegar a cotação', anchor='w')
labelSelecionarDia.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

calendarioMoeda = DateEntry(year=2024, locale='pt_br')
calendarioMoeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

labelTextoCotacao = tk.Label(text='')
labelTextoCotacao.grid(row=3, column=0,columnspan=2, padx=10, pady=10, sticky='nswe')

botaoPegarCotacao = tk.Button(text='Pegar cotação', command=pegarCotacao)
botaoPegarCotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nswe')

# Cotação varias moedas

labelCotacaoVariasMoedas = tk.Label(text='Cotação de múltiplas moedas', borderwidth=2, relief='solid')
labelCotacaoVariasMoedas.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

varCaminho = tk.StringVar()

labelSelecionarArquivo = tk.Label(text='Selecione um arquivo em excel:')
labelSelecionarArquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nswe')

botaoSelecionarArquivo = tk.Button(text='Clique para selecionar', command=selecionarArquivo)
botaoSelecionarArquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nswe')

labelArquivoSelecionado = tk.Label(text='Nenhum arquivo selecionado', anchor='e')
labelArquivoSelecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nswe')

labelDataInicial = tk.Label(text='Data inicial', anchor='w')
labelDataInicial.grid(row=7,column=0, padx=10, pady=10, sticky='nswe')
labelDataFinal = tk.Label(text='Data Final', anchor='w')
labelDataFinal.grid(row=8,column=0, padx=10, pady=10, sticky='nswe')

calendarioDataInicial = DateEntry(year=2024, locale='pt_br')
calendarioDataInicial.grid(row=7,column=1, padx=10, pady=10, sticky='nswe')
calendarioDataFinal = DateEntry(year=2024, locale='pt_br')
calendarioDataFinal.grid(row=8,column=1, padx=10, pady=10, sticky='nswe')

botaoAtualizarCotacao = tk.Button(text='Atualizar cotação', command=atualizarCotacao)
botaoAtualizarCotacao.grid(row=9, column=0, padx=10, pady=10, sticky='nswe')

labelAtualizarCotacao = tk.Label(text='')
labelAtualizarCotacao.grid(row=9, column=1,columnspan=2,padx=10, pady=10, sticky='nswe')

botaoFechar = tk.Button(text='Fechar', command=janela.quit)
botaoFechar.grid(row=10, column=2, padx=10, pady=10, sticky='nswe')

janela.mainloop()