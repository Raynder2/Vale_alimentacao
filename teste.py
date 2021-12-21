from tkinter import *
import clipboard as cb
import tkinter.font as font
import pyautogui as pg
from tkinter import messagebox
from datetime import date
import pandas as pd
from openpyxl import *
import time
import sys
import os
import re

cont = 2
contator = 0

meses = ['Janeiro','Fevereiro', 'Marco','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
data_atual = date.today()
pasta = "K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/5-Auxilio Alimentação/"+str(data_atual.month)+"-"+meses[data_atual.month-1]+" "+str(data_atual.year)+'/ORIGINAIS UNIDADES'
arquivos = os.listdir(pasta)

lotacao = pd.read_excel("K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/5-Auxilio Alimentação/1-PLANILHA PARA CORRIGIR LOTAÇÃO.xlsx")
lotacoes = pd.DataFrame(lotacao, columns = ['Cód. Departamento 2', 'Nome do Departamento'])

for arq in arquivos:
    # messagebox.askquestion("Iniciar Automoção", "Para evitar erros, garanta que a coluna valor esteja formatada como numeros.",icon ='info') == 'yes' and
    if(str(arq) != 'Thumbs.db'):
        data_atual = date.today()

        pasta = "K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/5-Auxilio Alimentação/"+str(data_atual.month)+"-"+meses[data_atual.month-1]+" "+str(data_atual.year)+'/ORIGINAIS UNIDADES'

        if os.path.exists(pasta):
            planilha = pasta+"/"+arq
            if os.path.isfile(planilha):
                valor = ''
                cont2 = ''

                valor2 = ''
                # Carregando planilha com as informações
                df = pd.read_excel(planilha)

                # print(df)

                for index, row in df.iterrows():
                    if("MATRÍCULA" in str(row) or "NOME DOS PROFISSIONAIS" in str(row)):
                        valor = str(valor)+str(index)+','
                        cont2 = index

                    limite = index

                # valor = valor.split(',')

                i = 0
                drop = 0

                for i in range(cont2):
                    if(i != valor):
                        df = df.drop(i)

                for index, row in df.iterrows():
                    if("LEGENDAS" in str(row)):
                        drop = 1
                    if(drop == 1):
                        df = df.drop(int(index))

                cod = arq.split('-')[0]
                df.rename({"Unnamed: 7": "Cód. Departamento 2"}, axis=1, inplace=True)
                df = pd.DataFrame( df, columns=['Unnamed: 0','Unnamed: 1','Unnamed: 2','Unnamed: 3','Unnamed: 4','Unnamed: 5','Unnamed: 6','Cód. Departamento 2','Unnamed: 8','Unnamed: 9','Unnamed: 1','Unnamed: 1','Unnamed: 1','Unnamed: 1','Unnamed: 1','Unnamed: 1','Unnamed: 1','Unnamed: 1','Unnamed: 1','Unnamed: 1','Unnamed: 2','Unnamed: 2','Unnamed: 2','Unnamed: 2','Unnamed: 2','Unnamed: 2','Unnamed: 2','Unnamed: 2','Unnamed: 2','Unnamed: 2','Unnamed: 3','Unnamed: 3','Unnamed: 3','Unnamed: 3','Unnamed: 3','Unnamed: 3','Unnamed: 3','Unnamed: 3','Unnamed: 3','Unnamed: 3','Unnamed: 4','Unnamed: 4','Unnamed: 4','Unnamed: 4','Unnamed: 4','Unnamed: 4'])
                df['Cód. Departamento 2'] = cod


                print(df.loc[df["Cód. Departamento 2"] == cod])
                # print(lotacoes["Cód. Departamento 2"])

                # resul = pd.merge(df['Cód. Departamento 2'], lotacoes, on=['Cód. Departamento 2'], how='left')
                # lotac = resul.iloc[2]['Nome do Departamento']

                # print(resul)

                if(contator == 0):
                    dfaux = df
                else:
                    dfaux = pd.concat([dfaux, df])

                df.to_excel('fim'+str(contator)+'.xlsx')
                contator = contator+1
        else:
            messagebox.showwarning("Planilha não encontrada", "Verifique se foi criada a pasta ORIGINAIS UNIDADES.")
    else:
        messagebox.showwarning("Pasta não encontrada", "Verifique o nome da pasta do mes atual se esta conforme o padrão: Relatorio_[MES NUMERO]-[MES POR EXTENSO] [ANO]'.xlsx")


dfaux = dfaux.loc[dfaux['Unnamed: 2'] != 0]
dfaux.dropna(subset=['Unnamed: 2'], inplace=True)
dfaux = dfaux.loc[dfaux['Unnamed: 1'] != 'MATRÍCULA']

dfaux.rename({"Unnamed: 0": "Nº","Unnamed: 1": "MATRÍCULA","Unnamed: 2": "NOME DOS PROFISSIONAIS","Unnamed: 6": "LOCAL","Unnamed: 7": "Cód. Departamento 2","Unnamed: 8": "FUNÇÃO"}, axis=1, inplace=True)
inicio = 0

for inicio in range(50):
    if(inicio != 0 and inicio != 1 and inicio != 2 and inicio != 8 and inicio != 7 and inicio != 6):
        dfaux.rename({"Unnamed: "+str(inicio): ""}, axis=1, inplace=True)


dfaux.to_excel('final.xlsx')