import wx
from openpyxl import Workbook
import v1
import os
import glob
#import conexao

# -*- coding: utf-8 -*-
class Main(wx.Frame):
    def __init__(self, title="Gerar OPs"):

        wx.Frame.__init__(self, None, title=title)

        panel = wx.Panel(self)
        
        texto = wx.StaticText(self, label=u"GDBR Toyoda Gosei", pos=(45, 90))
        texto.SetForegroundColour("white")
        
        btnGerar = wx.Button(self, label='Gerar', pos=(60, 50))

        self.Bind(wx.EVT_BUTTON, self.eventBtnGerar, btnGerar)

        self.SetBackgroundColour("#130f40")
        self.SetTitle('GDBR TOYODA')
        self.SetSize((200, 170))
        self.Centre()
        self.Show(True)




    def eventBtnGerar(self, event):
        i = 0
        arquivo = []
        informaData = wx.TextEntryDialog(self, 'Informe a data: XX/XX/XXXX ')
        
        if informaData.ShowModal() == wx.ID_OK:
            data = str(informaData.GetValue())
        informaData.Destroy()
        
        while i < 5:
            if i == 0:
                caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\MIX\\MONTH PLANNING\\FILTRAGEM'
            elif i == 1:
                caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\MIX\\MONTH PLANNING\\PELLET'
            elif i == 2:
                caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\MIX\\MONTH PLANNING\\ROLO'
            elif i == 3:
                caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\EXTRUSÃO\\SPONGE LINE\\MONTH PLANNING'
            elif i == 4:
                caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\TPV LINE\\MONTH PLAN'

            
            os.chdir(caminho)

            for file in glob.glob(data[3:5]+ '*'):
                arquivo.append(caminho + '\\' + file)

            i = i + 1
        v1.v01(arquivo,data)
        wx.MessageBox('Tudo Pronto!', 'Info', wx.OK | wx.ICON_INFORMATION)
         
        

if __name__ == "__main__":
    ex = wx.App(False)
    Main()
    ex.MainLoop()



