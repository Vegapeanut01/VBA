import win32com.client
import sys
import subprocess
import time 

class SapGui():
    def __init__(self):
        self.path = r"Caminho  do exe"
        subprocess.Popen(self,path)
        time.sleep(5)

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto)== win32com.client.CDispatch:
            return

        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection("Conexão", true)
        time.sleep(3)
        self.session = self.connection.Children(0)
        sefl.session.findById("wnd[0]").maximize
    
    def sapLogin(self):

if __name__ == '__main__':
    a = SapGui()
    a.__init__()