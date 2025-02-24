# -*- coding: utf-8 -*-
import wx
from openpyxl import Workbook
import v12

# -*- coding: utf-8 -*-
class Main (wx.Frame):
        def __init__(self, title="Apontamentos"):
            wx.Frame.__init__(self, None, title=title)

            panel = wx.Panel(self)
            
            meutexto = wx.StaticText(self, label=u"GDBR Toyoda Gosei", pos=(45, 90))
            meutexto.SetForegroundColour("white")
    

            btnAP = wx.Button(self, label = 'AP', pos=(60,50))



            self.Bind( wx.EVT_BUTTON, self.eventBtnAP, btnAP)
            
            self.SetBackgroundColour("#130f40")
            self.SetTitle('GDBR TOYODA')
            self.SetSize((200,170))
            self.Centre()
            self.Show(True)
        
        

        def eventBtnAP(self, event): 
            with wx.FileDialog(self, "Open XYZ file", wildcard="Excel files (*.xlsx)|*.xlsx", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
                if fileDialog.ShowModal() == wx.ID_CANCEL:
                    return     # the user changed their mind"
                # Proceed loading the file chosen by the user
                pathname = fileDialog.GetPath()
                v12.v12(pathname)
                wx.MessageBox('Tudo Pronto!', 'Info', wx.OK | wx.ICON_INFORMATION)


if __name__ == "__main__":
   ex = wx.App()
   Main()
   ex.MainLoop()