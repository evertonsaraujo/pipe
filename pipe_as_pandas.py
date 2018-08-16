import pandas
#import numpy as np
#import matplotlib.pyplot as plt
import pandas as pd
#import datetime
import os
from pandas import ExcelWriter
from pandas import ExcelFile
import xlsxwriter

#VARIAVEIS GLOBAIS

# REALIZAR BACKUP DAS PLANILHAS

# Salvando backup pipeline
#pipearq = '"Pipeline HANDEL.xlsx"'
#pipebkp = '"Pipeline HANDEL-bkp.xlsx"'
#pipefile = os.system('cp ' + pipearq + ' ' + pipebkp)
#os.system('mkdir backup')
#os.system('mv ' + pipebkp + ' backup')

# ABRE PLANILHA GSA
gsa = pd.ExcelFile('Status GSA Orders 23.05.xlsx')
gsadf = pd.read_excel(gsa, '2018')
gsadf.to_excel('Status GSA Orders 23.05.xlsx', sheet_name='2018')

# ABRE PLANILHA TEXTILE
#tex = pd.ExcelFile('Textile Order Status 28.05.xlsx')
#texdf = pd.read_excel(tex,'2018')
#texdf.to_excel('Textile Order Status 28.05.xlsx', sheet_name='2018') # Planilha muito pesada
#texdf.dropna(['IMP NR.'])

# ABRE PLANILHA PIPELINE
pipeline = pd.ExcelFile('Pipeline HANDEL.xlsx')
pipedf = {sheet_name: pipeline.parse(sheet_name)
          for sheet_name in pipeline.sheet_names}
pipepipe = pd.read_excel(pipeline,sheet_name='PIPELINE')

# ESCREVE NOVA PLANILHA PIPELINE
dfpipenew = pd.DataFrame.copy(pipepipe)
pipenew = ExcelWriter('Pipelinenew.xlsx', engine='xlsxwriter')
dfpipenew.to_excel(pipenew, sheet_name='PIPELINE')
#pipenew.save()
#print(dfpipenew.head())


for index, row in pipepipe.iterrows():
    for index, rowgsa in gsadf.iterrows():
        if row['IMP'] == rowgsa['IMP NR.']:
            if pandas.isna(rowgsa['CURRENT\nETD']):
                print('Pipeline IMP: ', row['IMP'], ' - ', 'GSA IMP: ', rowgsa['IMP NR.'], ' -- DATA NAO INFORMADA\n')
            elif rowgsa['CURRENT\nETD'] > row['ETD_PREVISTO']:
                    print('Pipeline IMP: ', row['IMP'], row['ETD_PREVISTO'], '--', 'IMP GSA', rowgsa['IMP NR.'], rowgsa['CURRENT\nETD'], ' -- Data Desatualizada\n')
                    # Atualizar data em Planilha Pipeline
            elif rowgsa['CURRENT\nETD'] < row['ETD_PREVISTO']:
                    print('Pipeline IMP: ', row['IMP'], ' - ', 'GSA IMP: ', rowgsa['IMP NR.'], rowgsa['CURRENT\nETD'], ' -- ATENCAO - DATA MAIOR EM PIPELINE\n')
            elif rowgsa['CURRENT\nETD'] == row['ETD_PREVISTO']:
                    print('Pipeline IMP: ', row['IMP'], row['ETD_PREVISTO'], ' - ', 'GSA IMP: ', rowgsa['IMP NR.'], rowgsa['CURRENT\nETD'], ' -- DATA PIPELINE ATUALIZADA.. TUDO OK!\n')
pipenew.save()
pipenew.close()