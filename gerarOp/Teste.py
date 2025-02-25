# -*- coding: utf-8 -*-
from numpy.core.numeric import zeros_like
import openpyxl as xl
import wx
from openpyxl import Workbook
import datetime
import os
import glob
import criar_xlsx
from decimal import *
from conexao import *
import numpy
import json

valores = '2'

op = open("C:\TOTVS\op.txt", "w")
op.write("1" + valores + "\n")
op.close