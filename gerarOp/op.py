import wx
from openpyxl import Workbook
import v1
import os
import glob
import datetime

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
        arquivo     = []
        informaData = wx.TextEntryDialog(self, 'Informe a data: XX/XX/XXXX ')
        i = 0
        
        if informaData.ShowModal() == wx.ID_OK:
            data = str(informaData.GetValue())
            nOP = data[8:10] + data[3:5] + data[0:2]  

        informaData.Destroy()

        log = open("C:\TOTVS\log" + nOP + ".txt", "w")

        #se a data não estiver no formato correto 
        try:
            datetime.datetime.strptime(data, "%d/%m/%Y")
        except ValueError:
                wx.MessageBox('Por favor, informe uma data valida, no formato dd/mm/aaaa', 'Data invalida!', wx.ICON_INFORMATION)
        else:
        
            while i < 5:
                if i == 0:
                    caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\SPONGE LINE\\MONTH PLANNING'
                elif i == 1:
                    caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\MIX\\MONTH PLANNING\\PELLET'
                elif i == 2:
                    caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\MIX\\MONTH PLANNING\\ROLO'
                elif i == 3:
                     caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\MIX\\MONTH PLANNING\\FILTRAGEM'
                elif i == 4:
                    caminho = '\\\\files-gdbr01\\gdbr\\GeDoc\\GeDoc - Public\\Outros\\Production\\1 - PLANEJAMENTO DE PRODUÇÃO - Production planning\\4 - WS\\EXTRUSÃO\\TPV LINE\\MONTH PLAN'

                
                os.chdir(caminho)

                encontrouArquivo = False

                for file in glob.glob(data[3:5]+ '*'):
                    arquivo.append(caminho + '\\' + file)
                    encontrouArquivo = True

                #planilha não foi encontrada ou houve erro ao abrir o arquivo
                if encontrouArquivo == False:
                    log.write(f"[{datetime.datetime.now()}] Erro: Nenhum arquivo encontrado no caminho {caminho}\n")
                    log.write(f"[{datetime.datetime.now()}] Erro: Falha ao abrir o arquivo no caminho {caminho}\n")

                i = i + 1

            self.show_loading()
            try:
                v1.v01(arquivo, data)
            finally:
                self.hide_loading()

            log.close()

            log_path = "C:\TOTVS\log" + nOP + ".txt"

            if os.path.exists(log_path) and os.stat(log_path).st_size > 0:
                wx.MessageBox('Falha ao gerar OPs! Verifique o arquivo de log', 'Info', wx.OK | wx.ICON_INFORMATION)
                os.startfile(log_path)
            else:
                wx.MessageBox('OPs geradas com sucesso!', 'Info', wx.OK | wx.ICON_INFORMATION)
                self.Close()

    # Tela de carregamento
    def show_loading(self):
        self.loading_frame = wx.Frame(None, title="Carregando", size=(200, 100))
        panel = wx.Panel(self.loading_frame)
        wx.StaticText(panel, label="Por favor, aguarde...", pos=(40, 20))
        self.loading_frame.Centre()
        self.loading_frame.Show()
        wx.GetApp().Yield() 

    def hide_loading(self):
        self.loading_frame.Destroy()

if __name__ == "__main__":
    ex = wx.App(False)
    Main()
    ex.MainLoop()