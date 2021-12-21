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

meses = ['Janeiro','Fevereiro', 'Marco','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

##if(messagebox.askquestion("Iniciar Automoção", "Para evitar erros, garanta que a coluna valor esteja formatada como numeros.",icon ='info') == 'yes'):
data_atual = date.today()

pasta = "K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/5-Auxilio Alimentação/"+str(data_atual.month)+"-"+meses[data_atual.month-1]+" "+str(data_atual.year)
planilha = 'Relatorio_'+meses[data_atual.month-1]+str(data_atual.year)+'.xlsx'


   if os.path.exists(pasta):
       planilha = pasta+"/"+planilha
       if os.path.isfile(planilha):
           # Carregando planilha com as informações
           df = pd.read_excel(planilha)

           # PEGANDO TODAS OS DATAFRAMES QUE SERÃO NECESSARIOS PARA ANALISE DE DADOS
           geral = pd.read_excel("K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/1-RELATORIO GERAL MENSAL/"+str(data_atual.month-1)+"-SMS GERAL "+str(data_atual.month-1)+"."+str(data_atual.year)+".xlsx")
           rgs = pd.DataFrame(geral, columns = ['CPF', 'RG'])
           cargos = pd.DataFrame(geral, columns = ['Matricula','Atividade','Cargo'])
           pagamentos = pd.DataFrame(geral, columns = ['Matricula','V. Alim. Ag','V. Alim. Mot'])

           lotacao = pd.read_excel("K:/Administrativo/SetorPessoal/Marcelo/1-Arquivos SMS/5-Auxilio Alimentação/1-PLANILHA PARA CORRIGIR LOTAÇÃO.xlsx")
           codigos = pd.DataFrame(lotacao, columns = ['Nome Departamento', 'Cód. Departamento 2'])
           lotacoes = pd.DataFrame(lotacao, columns = ['Codigo Departamento', 'Nome do Departamento'])

           #LIMPAR O NOME DA MAE
           df.loc[df['Nome']!='0','Nome Mae']=''

           # ESCLUIR ZERADOS
           df = df.loc[df['Valor'] != '0,00']
           df = df.loc[df['Valor'] != '0']
           df = df.loc[df['Valor'] != 0]

           # CRIAR NOVAS COLUNAS
           # df.insert(5, "ATIVIDADE", "", allow_duplicates=False)

           #APAGAR RG
           df = df.drop(columns='RG')

           # TRAZER NOVOS RGS
           df = pd.merge(df, rgs.drop_duplicates(), on=['CPF'], how='left')

           # TRAZER ATIVIDADE E CARGO
           df = pd.merge(df, cargos.drop_duplicates(), on=['Matricula'], how='left')

           # ORGANIZAR PLANILHA
           df = pd.DataFrame( df, columns=['Matricula','Nome','CPF','RG','Data Nascimento','Atividade','Cargo','Nome Mae','Endereco','Numero','Complemento','Bairro','CEP','Nome Departamento','Codigo Departamento','UF','Valor','Numero Registro'])

           # REMOVER MOTORISTAS E AGENTES DE ENDEMIAS QUE NÃO SAO READAPTADOS
           removidos = len(df.index)
           filtro = df[ (df['Atividade'] != 'Readaptado-servicos Diversos') & (df['Cargo'] == 'Motorista') ].index
           filtro2 = df[ (df['Atividade'] != 'Readaptado-servicos Diversos') & (df['Cargo'] == 'Agente de Combate As Endemias') ].index
           df.drop(filtro , inplace=True)
           df.drop(filtro2 , inplace=True)

           # CONTANDO SERVIDORES REMOVIDOS POR CARGO E ATIVIDADE
           removidos2 = len(df.index)
           removidos = removidos - removidos2

           # REMOVER MOTORISTAS E AGENTES DE ENDEMIAS QUE RECEBERAM NA FOLHA
           # VERIFICAR NA PLANILHA REGISTRO GERAL POR MATRICULA

           #APAGAR Cod Departamento
           df = df.drop(columns='Codigo Departamento')
           df = pd.merge(df, codigos.drop_duplicates(), on=['Nome Departamento'], how='left')
           df = df.rename(columns={'Cód. Departamento 2': 'Codigo Departamento'})

           df = df.drop(columns='Nome Departamento')
           df = pd.merge(df, lotacoes.drop_duplicates(), on=['Codigo Departamento'], how='left')

           # ORGANIZAR PLANILHA
           df = pd.DataFrame( df, columns=['Matricula','Nome','CPF','RG','Data Nascimento','Atividade','Cargo','Nome Mae','Endereco','Numero','Complemento','Bairro','CEP','Nome do Departamento','Codigo Departamento','UF','Valor','Numero Registro'])

           print(pagamentos)
           # ESCLUIR LOTAÇÕES SEM DIREITO
           rem = len(df.index)
           df = df.loc[df['Codigo Departamento'] != 'NÃO RECEBEM']
           rem = rem - len(df.index)

           duploV = len(df.index)
           df = df.groupby('CPF').agg({
                   'Matricula': 'first',
                   'Nome': 'first',
                   'CPF': 'first',
                   'RG': 'first',
                   'Data Nascimento': 'first',
                   'Atividade': 'first',
                   'Cargo': 'first',
                   'Nome Mae': 'first',
                   'Endereco': 'first',
                   'Numero': 'first',
                   'Complemento': 'first',
                   'Bairro': 'first',
                   'CEP': 'first',
                   'Nome do Departamento': 'first',
                   'Codigo Departamento': 'first',
                   'UF': 'first',
                   'Valor': sum,
                   'Numero Registro': 'first'
               })
           duploV = duploV - len(df.index)

           # TRAZER PAGAMENTOS
           df = pd.merge(df, pagamentos, on=['Matricula'], how='left')

           # df = df.loc[df['Codigo Departamento'] != 'NÃO RECEBEM']
           rem2 = len(df.index)
        #    Olhar no Init para saber como pegar valores corretos nessa parte
        #    df = df.loc[df['V. Alim. Ag'].isnull()]
        #    df = df.loc[df['V. Alim. Mot'].isnull()]
           rem2 = rem2 - len(df.index)

           df = df.drop(columns="V. Alim. Ag")
           df = df.drop(columns="V. Alim. Mot")

           finalPlanilha = pasta+'/1-FREQUENCIA SAUDE '+str(data_atual.month)+'.'+str(data_atual.year)+'.xlsx'
           if os.path.isfile(finalPlanilha):
               finalPlanilha2 = pasta+'/1-FREQUENCIA SAUDE '+str(data_atual.month)+'.'+str(data_atual.year)+' V2.xlsx'
               if os.path.isfile(finalPlanilha2):
                   df.to_excel(pasta+'/1-FREQUENCIA SAUDE '+str(data_atual.month)+'.'+str(data_atual.year)+' V3.xlsx')
               else:
                   df.to_excel(finalPlanilha2)
           else:
               df.to_excel(finalPlanilha)

           relatorioFinal = "Resultado:\n\nZerados excluidos.\nRgs corrigidos.\nAtividades e Cargos puxados.\n"+str(removidos)+" servidores removidos por cargo e atividade.\nLotações e codigos corrigidos.\n"+str(rem)+" servidores removidos por lotação.\n"+str(duploV)+" Servidores estavão duplicados, os valores foram somados e removido a duplicidade\nForam encontrados "+str(rem2)+" motoristas ou agentes readaptados que ja recebem em folha, os mesmos foram removidos."
           messagebox.showinfo(title="Concluido", message=relatorioFinal)
           #fim
       else:
           messagebox.showwarning("Planilha não encontrada", "Verifique se foi baixado a o relatorio em https://outprod01.goiania.go.gov.br/frequenciasaude/Relatorios.aspx.")
   else:
       messagebox.showwarning("Pasta não encontrada", "Verifique o nome da pasta do mes atual se esta conforme o padrão: Relatorio_[MES NUMERO]-[MES POR EXTENSO] [ANO]'.xlsx")