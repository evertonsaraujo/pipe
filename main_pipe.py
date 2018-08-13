#!/usr/bin/python
# Importando modulo de leitura de planilhas
import datetime
import xlrd
import xlsxwriter
import sys
import os
# ----------------------------------INICIO DO CODIGO DE LEITURA DO PIPELINE -----------------------------------------#


print('Iniciando leitura das Planilhas PIPELINE, GSA e Textile....\n\n')

# Abre arquivo Workbook PIPELINE, GSA e Textile
wbpippeline = xlrd.open_workbook('Pipeline HANDEL.xlsx')
wbgsa = xlrd.open_workbook('Status GSA Orders 23.05.xlsx')
wbtex = xlrd.open_workbook('Textile Order Status 28.05.xlsx')


# Abre Worksheet Planilha PIPELINE, GSA e Textile
print('Abrindo pastas das planilhas..... \n\n')
wspipeline = wbpippeline.sheet_by_name('PIPELINE')
wsgsa = wbgsa.sheet_by_name('2018')
wstex = wbtex.sheet_by_name('2018')

# Corpiar Pipeline Handel para Pipeline Handel Atualizada


# Variaveis Globais

total_rowspipe = wspipeline.nrows
total_colspipe = wspipeline.ncols
total_rowsgsa = wsgsa.nrows
total_colgsa = wsgsa.ncols
total_rowstex = wstex.ncols
total_coltex = wstex.ncols

# Compara a lista pipeline e GSA e informa se as datas estão atualizadas.

print('Procurando NRImp de PIPELINE em Planilha GSA... \n')
for x in range(total_rowspipe):
    for y in range(total_rowsgsa):
        if wspipeline.cell_value(x, 0) == wsgsa.cell_value(y, 0):
            py, pM, pd, ph, pm, ps = xlrd.xldate_as_tuple(wspipeline.cell_value(x, 17), wbpippeline.datemode)
            pipeDate = datetime.date(py, pM, pd)
            if wsgsa.cell_value(y, 35) == '':
                print(wsgsa.cell_value(y, 0) + ' ------  Data ETD PREVISTO GSA não informada \n')
                y = + 1
            else:
                gy, gM, gd, gh, gm, gs = xlrd.xldate_as_tuple(wsgsa.cell_value(y, 35), wbgsa.datemode)
                print(wspipeline.cell_value(x, 0) + ' ' + "{0}/{1}/{2}".format(pd, pM, py) + ' ' + wsgsa.cell_value(y, 0) + ' ' + "{0}/{1}/{2}".format(gd, gM, gy))
                gsaDate = datetime.date(gy, gM, gd)
                if gsaDate == pipeDate:
                    print('Data Atualizada -- Tudo ok \n')
                elif (gsaDate > pipeDate):
                    print('ATENCAO --- Data desatualizada \n')
                else:
                    print('ATENCAO --- Data do Pipeline Maior \n')
        else:
            y += 1
    else:
        x += 1

# Salvando backup pipeline

pipefile = os.system('cp ' ' Pipeline HANDEL.xls ''./'' "Pipeline HANDEL.bkp.xlsx"')

# Abrindo planilha para salvar dados
#wbwrpipe = xlsxwriter.Workbook('Pipeline HANDEL.xlsx')
#wswrpipe = wbwrpipe.add_worksheet('PIPELINE')


# Compara a lista pipeline e Textile e informa se as datas estão atualizadas.
print('Procurando NRImp de PIPELINE em Planilha TEX... \n')
for xx in range(total_rowspipe):
    for yy in range(total_rowstex):
        if wspipeline.cell_value(xx, 0) == wstex.cell_value(yy, 0):
            py, pM, pd, ph, pm, ps = xlrd.xldate_as_tuple(wspipeline.cell_value(xx, 17), wbpippeline.datemode)
            pipeDate = datetime.date(py, pM, pd)
            if wstex.cell_value(yy, 38) == '':
                print(wstex.cell_value(yy, 0) + ' ------  Data ETD PREVISTO TEX não informada \n')
                yy = + 1
            else:
                ty, tM, td, th, tm, ts = xlrd.xldate_as_tuple(wstex.cell_value(yy, 38), wbtex.datemode)
                print('NrIMP Pipeline :' + wspipeline.cell_value(xx, 0) + ' ' + "{0}/{1}/{2}".format(pd, pM, py) + ' ; NrIMP TEX :' + wstex.cell_value(yy, 0) + ' ' + "{0}/{1}/{2}".format(td, tM, ty))
                texDate = datetime.date(ty, tM, td)
                if texDate == pipeDate:
                    print('Data Atualizada -- Tudo ok \n')
                elif texDate < pipeDate:
                    print('ATENCAO --- Data Pipeline Maior que TEX\n')
                else:
                    print('ATENCAO --- Data desatualizada... Atualizando....\n ')
                    #wswrpipe.a(xx, 17, texDate)


        else:
            yy += 1
    else:
        xx += 1

