# -*- coding: utf-8 -*-
import openpyxl
from openpyxl import Workbook
import datetime
import xlsxwriter
import wx
import os


def xlsx(arqNome, pathName):
    path = pathName
    nomeArq = arqNome
    

    workbook = xlsxwriter.Workbook(path+'\\AP - '+nomeArq+'.xlsx')
    worksheet = workbook.add_worksheet()   
    
    workbook.close()
    


