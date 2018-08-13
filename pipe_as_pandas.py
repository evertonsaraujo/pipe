import pandas
import numpy
import matplotlib.pyplot as plt
import pandas as pd
import datetime
import os

#VARIAVEIS GLOBAIS

# REALIZAR BACKUP DAS PLANILHAS


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
pipedf = pd.read_excel(pipeline, 'PIPELINE')
pipedf.to_excel('Pipeline HANDEL.xlsx', sheet_name='PIPELINE')
for index, row in pipedf.iterrows():
    for index, rowgsa in gsadf.iterrows():
        if row['IMP'] == rowgsa['IMP NR.']:
            if pandas.isna(rowgsa['CURRENT\nETD']):
                print('Pipeline IMP: ', row['IMP'], ' - ', 'GSA IMP: ', rowgsa['IMP NR.'], ' -- DATA NAO INFORMADA\n')
            elif rowgsa['CURRENT\nETD'] > row['ETD_PREVISTO']:
                    print('IMP Pipeline', row['IMP'], row['ETD_PREVISTO'], '--', 'IMP GSA', rowgsa['IMP NR.'], rowgsa['CURRENT\nETD'], ' -- Data Desatualizada\n')
                    #pipedf.apply(func=setattr()) ESTUDANDO AQUI...
            elif rowgsa['CURRENT\nETD'] < row['ETD_PREVISTO']:
                    print('Pipeline IMP: ', row['IMP'], ' - ', 'GSA IMP: ', rowgsa['IMP NR.'], rowgsa['CURRENT\nETD'], ' -- ATENCAO - DATA MAIOR EM PIPELINE\n')
            elif rowgsa['CURRENT\nETD'] == row['ETD_PREVISTO']:
                    print('Pipeline IMP: ', row['IMP'], row['ETD_PREVISTO'], ' - ', 'GSA IMP: ', rowgsa['IMP NR.'], rowgsa['CURRENT\nETD'], ' -- DATA PIPELINE ATUALIZADA.. TUDO OK!\n')
