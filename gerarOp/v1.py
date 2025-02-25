# -*- coding: utf-8 -*-
from numpy.core.numeric import zeros_like
import openpyxl as xl
import wx
from openpyxl import Workbook
import datetime
import os
import criar_xlsx
from decimal import *
from conexao import *
import numpy
import json

def v01(path,data):
    pathAp  = path
    dataAp  = data
    tpAp    = xl.load_workbook(pathAp, data_only=True)
    tabAp   = tpAp.active
    nOP = ''
    i = 2
    j = 5
    valor = []


    cursor = dbConn()
    cursor.execute("SELECT DISTINCT ZKI_KANBAN 'Kanban', RTRIM(CAD.[PN]) 'PN', ZKI_ITEMOP, SB1.B1_PESO, IIF(SB1.B1_PESO > 0, CAST(((1 / SB1.B1_PESO)) AS NUMERIC(16,6)), 0) 'Peca por KG', RTRIM(CAD.[LINHA]) 'LINHA', rtrim(SB1.B1_LOCPAD) 'Armazem' FROM ZKI010 ZKI WITH (NOLOCK) LEFT JOIN (SELECT CASE ISNULL(ZKB.ZKB_KANBAN,'') WHEN '' THEN B1_ZZKNBAN ELSE ZKB.ZKB_KANBAN END [KANBAN], B1_COD [PN], B1_ZZLNPRD [LINHA] FROM SB1010 (NOLOCK) SB1 LEFT JOIN ZKB010 ZKB WITH (NOLOCK) ON SB1.B1_COD = ZKB.ZKB_PN AND ZKB.D_E_L_E_T_ = '' WHERE SB1.D_E_L_E_T_ = '' AND SB1.B1_ZZKNBAN <> '') [CAD] ON CAD.KANBAN = ZKI.ZKI_KANBAN INNER JOIN SB1010 SB1 WITH (NOLOCK) ON SB1.B1_COD = CAD.[PN] WHERE ZKI.D_E_L_E_T_ = '' AND ZKI.ZKI_STATUS = 'L';")
    
    vetDados = cursor.fetchall()

    wb1 = xl.load_workbook(pathAp, data_only = True)

    ws1 = wb1['RESUMO']

    while i < 28:
        
        if ws1.cell(3,i).value != None:
            if  data[0:5] in str(format(ws1.cell(3,i).value, "%d/%m")):

                while j < 11:     

                    if ws1.cell(j,i).value != '' and ws1.cell(j,i).value != 0:
                        
                        certo = (str(ws1.cell(j,2).value) + ' , ' + str(ws1.cell(j,i).value))
                        valor.append(certo)

                    j = j + 1
        i = i + 1

    print(valor)
    return 