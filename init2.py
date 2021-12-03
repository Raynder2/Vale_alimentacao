from tkinter import *
import clipboard as cb
import tkinter.font as font
import pyautogui as pg
from tkinter import messagebox
from datetime import date
import pandas as pd
from openpyxl import *
import shutil #Mover arquivos
import time
import sys
import os
import re

cont = 2
contator = 0
barraD = 0
barraN = 0
moves = [0]*20
meses = ['Janeiro','Fevereiro', 'Marco','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
data_atual = date.today()
pasta = "K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/5-Auxilio Alimentação/"+str(data_atual.month)+"-"+meses[data_atual.month-1]+" "+str(data_atual.year)+'/ORIGINAIS UNIDADES'
arquivos = os.listdir(pasta)

lotacao = pd.read_excel("K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/5-Auxilio Alimentação/1-PLANILHA PARA CORRIGIR LOTAÇÃO.xlsx")
lotacoes = pd.DataFrame(lotacao, columns = ['Cód. Departamento 2', 'Nome do Departamento'])

for arq in arquivos:
    # messagebox.askquestion("Iniciar Automoção", "Para evitar erros, garanta que a coluna valor esteja formatada como numeros.",icon ='info') == 'yes' and 
    if(str(arq) != 'Thumbs.db' and str(arq) != 'LANÇADAS'):
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
                df['Unnamed: 7'] = cod
                df.rename({"Unnamed: 7": "Cód. Departamento 2"}, axis=1, inplace=True)

                lotacoes = lotacoes.loc[lotacoes['Cód. Departamento 2'] != 'NÃO RECEBEM']
                lotacoes['Cód. Departamento 2'].astype(int)
                lotac = lotacoes.loc[lotacoes['Cód. Departamento 2'] == int(cod)]
                lotac = lotac['Nome do Departamento'].values[0]

                df['Unnamed: 6'] = lotac
                
                if(contator == 0):
                    dfaux = df
                else:
                    dfaux = pd.concat([dfaux, df])

                contator = contator+1


                    # shutil.move(source,destination)


        else:
            messagebox.showwarning("Planilha não encontrada", "Verifique se foi baixado a o relatorio em https://outprod01.goiania.go.gov.br/frequenciasaude/Relatorios.aspx.")
    else:
        messagebox.showwarning("Pasta não encontrada", "Verifique o nome da pasta do mes atual se esta conforme o padrão: Relatorio_[MES NUMERO]-[MES POR EXTENSO] [ANO]'.xlsx")



dfaux = dfaux.loc[dfaux['Unnamed: 2'] != 0]
dfaux.dropna(subset=['Unnamed: 14'], inplace=True)
dfaux = dfaux.loc[dfaux['Unnamed: 1'] != 'MATRÍCULA']

dfaux.rename({"Unnamed: 0": "Nº","Unnamed: 1": "MATRÍCULA","Unnamed: 2": "NOME DOS PROFISSIONAIS","Unnamed: 6": "LOCAL","Unnamed: 7": "Cód. Departamento 2","Unnamed: 8": "FUNÇÃO"}, axis=1, inplace=True)


dfaux.reset_index(inplace=True, drop=False)

for index, row in dfaux.iterrows():
    barraD = len(str(row).split('/D')) + len(str(row).split('SD'))
    barraN = len(str(row).split('/N')) + len(str(row).split('SN'))

    dfaux.at[index,'Unnamed: 5'] = barraD - 2
    dfaux.at[index,'Unnamed: 4'] = barraN - 2
    
    barraD = 0
    barraN = 0

inicio = 0
for inicio in range(50): #41
    if(inicio != 0 and inicio != 1 and inicio != 2 and inicio != 8 and inicio != 7 and inicio != 6 and inicio != 5 and inicio != 4):
        dfaux.rename({"Unnamed: "+str(inicio): ""}, axis=1, inplace=True)

dfaux.rename({"Unnamed: 5": "/D","Unnamed: 4": "/N"}, axis=1, inplace=True)

dfaux = dfaux.loc[dfaux['MATRÍCULA'] != 'MATRÍCULA']
dfaux = dfaux.loc[dfaux['MATRÍCULA'] != 'MATRÍCULA/CPF']
dfaux = dfaux.loc[dfaux['MATRÍCULA'] != 'NOME DOS PROFISSIONAIS']

cont3 = 0
for index, row in dfaux.iterrows():
    codd = str(row).split('Cód. Departamento 2')[1]
    codd = codd.split('FUNÇÃO')[0]
    codd = codd.strip()
    if(codd not in moves):
        cont3 = cont3+1
        moves[cont3] = codd
        
for arq in arquivos:
    cod = arq.split('-')[0]
    if(cod in moves):
        source = pasta+"/"+arq
        destination = pasta+"/LANÇADAS/"+arq
        shutil.move(source,destination)

dfaux.to_excel('K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/5-Auxilio Alimentação/'+str(data_atual.month)+"-"+meses[data_atual.month-1]+" "+str(data_atual.year)+'/2-JUNÇÃO UNIDADES '+str(data_atual.month)+"."+str(data_atual.year)+'.xlsx')

# 27/setembro 
# brunna felipe
# 04564809199
# 26901247980

# for index, row in dfaux.iterrows():
#     codd = str(row).split('Cód. Departamento 2')[1]
#     codd = codd.split('FUNÇÃO')[0]
#     codd = codd.strip()
#     if(codd not in moves):
#         cont3 = cont3+1
#         moves[cont3] = codd
        
# for arq in arquivos:
#     cod = arq.split('-')[0]
#     if(cod in moves):
#         source = pasta+"/"+arq
#         destination = pasta+"/LANÇADAS/"+arq
#         shutil.move(source,destination)