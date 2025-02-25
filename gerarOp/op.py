import wx
from openpyxl import Workbook
import v1
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
        informaData = wx.TextEntryDialog(self, 'Informe a data: XX/XX/XXXX ')
        
        if informaData.ShowModal() == wx.ID_OK:
            data = str(informaData.GetValue())
        informaData.Destroy()

        pathname = "C:\\TOTVS\\" + data[3:5] + ' - PLANEJAMENTO SEMANAL - SUTORENA.xlsx'
        v1.v01(pathname,data)
        wx.MessageBox('Tudo Pronto!', 'Info', wx.OK | wx.ICON_INFORMATION)
         
        

if __name__ == "__main__":
    ex = wx.App(False)
    Main()
    ex.MainLoop()



