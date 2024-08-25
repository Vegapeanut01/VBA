'''
import win32com.client
import subprocess
import sys
import time 
import pandas as pd
'''
from tkinter import * #para o botão 
from tkinter import messagebox #para o botão 

'''
class SapGui(object): 
    def __init__(self):
        self.path #r caminho onde está intalado o sap .exe
        subprocess.Popen(self.path)


        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = self.SapGuiAuto.GetScriptingEngine

        #Aqui colocar o nome da Conexão do SAP
        #Na linha de baixo entre () colocar o nome da conexão 
        #botão direito (antes de entrar no 4hanna) e Caracteristicas e copiar o nome que está na descrição 
        #colar antes de True não esquecendo dos parâmetros
        self.connection = application.OpenConnection("BMF [ecc1.ddns.net]", True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize

    def sapLogin(self):
       try:
            #Client
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "800" 
            #User
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "usuario" 
            #Password
            self.session.findById("wnd[0]/usr/txtRSYST-BCODE").text = "senha" 
            #Logon Language
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "senha" 
            #Botão para executar
            self.session.findById("wnd[0]").sendVKey(0)

       except:
            print(sys.exc_info()[0])
       messagebox.showinfo("showinfo", "Login realizado com sucesso!")
       time.sleep(2) #depois do logion vai esperar dois segundos para executar o resto do código
       self.nome da função

    def pegarLote(self):
        #ler arquivo de uma pasta do excel
        data = pd.read_excel(caminho onde está a pasta "", sheet_name="nomeDaPlanilha").astype(str)
        data.columns = data.columns.str.replace(' ', '_')

        for index, row in data.iterrows():
            print(index, row.NOME)

        row.NOMEdaCOLUNA
        row.TERMO_DE_PESQUISA #vai com o _ porque o espaço foi substituído


    #O que o script gerar posso colocar dentro de uma função, mudando 
    #session por self.session e press por press()
    #sendVKey deve ser sendVKey(0)
'''
#Ate esse ponto já abre o SAP 
#Criando um botão 
if __name__ == '__main__':
    window = Tk()
    window.geometry("200x50")
    botao = Button(window, text="Login SAP", command= lambda :SapGui().sapLogin())
    botao.pack()
    mainloop()