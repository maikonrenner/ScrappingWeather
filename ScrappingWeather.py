import requests
import json
import time, sys
import os

#Frases que o bot reproduz durante a execução
frase = "Olá sou o CWBLab, "
frase2 = "Estou capturando informações do clima!\n"
frase3 = "Estou enviando os dados para os INDICADORES CLIMATICOS,"
frase4 = " Prontinho! ;)"

# Acessando o registro do windows para desativar o proxy corporativo
import win32com.client

try:
    import _winreg as winreg
except:
    import winreg

class HTTP:
    proxy = ""
    isProxy = False

    def __init__(self):
        self.get_proxy()

    def get_proxy(self):
        oReg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        oKey = winreg.OpenKey(oReg, r'Software\Microsoft\Windows\CurrentVersion\Internet Settings')
        dwValue = winreg.QueryValueEx(oKey, 'ProxyEnable')

        if dwValue[0] == 1:
            oKey = winreg.OpenKey(oReg, r'Software\Microsoft\Windows\CurrentVersion\Internet Settings')
            dwValue = winreg.QueryValueEx(oKey, 'ProxyServer')[0]
            self.isProxy = True
            self.proxy = dwValue

    def url_post(self, url, formData):
        httpCOM = win32com.client.Dispatch('Msxml2.ServerXMLHTTP.6.0')

        if self.isProxy:
            httpCOM.setProxy(2, self.proxy, '<local>')

        httpCOM.setOption(2, 13056)
        httpCOM.Open('POST', url, False)
        httpCOM.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded')
        httpCOM.setRequestHeader('User-Agent', 'whatever you want')

        httpCOM.send(formData)

        return httpCOM.responseText

http = HTTP()

for i in list(frase):
    print(i, end='')
    #O stdout só é atualizado quando há nova linha e como nós estamos mandando tudo na mesma é preciso forçar a atualização.
    sys.stdout.flush()
    time.sleep(0.05)


for i in list(frase2):
    print(i, end='')
    #O stdout só é atualizado quando há nova linha e como nós estamos mandando tudo na mesma é preciso forçar a atualização.
    sys.stdout.flush()
    time.sleep(0.05)

cwb = http.url_post('http://api.openweathermap.org/data/2.5/weather?id=6322752&lang=pt_br&units=metric&appid=ef4949b294870b8302cfcb0070703970', 'test=1')

#Convertendo para JSON o resultado do POST
cwb_weather = json.loads(cwb)

#Extraçao e tratamento dos dados capturados no requests
import pandas as pd 
from datetime import datetime #Chamada para a captura da data e hora
data_hora = datetime.today() #Data:Hora

# INICIO DOS DADOS CLIMATICOS DO PKB
cwb_table = list()
cwb_item = {

    "Periodo": data_hora,
    "Cidade": cwb_weather['name'],
    "Umidade": cwb_weather['main']['humidity'],
    #"Pressao": weather['main']['pressure'],
    "Temp_min": cwb_weather['main']['temp_min'],
    "Aparente": cwb_weather['main']['feels_like'],
    "Temp_max": cwb_weather ['main']['temp_max'],
    "Clima_principal": cwb_weather['weather'][0]['main'],
    "Descricao": cwb_weather['weather'][0]['description'],

}
cwb_table.append(cwb_item)
df_cwb = pd.DataFrame(cwb_table)
print(df_cwb)

# ----------------------------------------------------------------------------------------------------

#Acessando ao DB - SQL SERVER usando o pyodbc
import pyodbc

server = 'localhost' 
database = 'Clima' 
username = '' 
password = '' 
cnxn = pyodbc.connect('DRIVER={SQL SERVER Native Client 11.0};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

#inserindo o Dataframe no SQL Server:

for i in list(frase3):
    print(i, end='')
    #O stdout só é atualizado quando há nova linha e como nós estamos mandando tudo na mesma é preciso forçar a atualização.
    sys.stdout.flush()
    time.sleep(0.05)

for index, row in df_PKB.iterrows():
    cursor.execute("INSERT INTO Clima (Periodo,Cidade,Umidade,Temp_min,Aparente,Temp_max,Clima_principal,Descricao,Shopping) values(?,?,?,?,?,?,?,?,?)", row.Periodo, row.Cidade, row.Umidade,row.Temp_min, row.Aparente, row.Temp_max, row.Clima_principal, row.Descricao, row.Shopping)
cnxn.commit()

cursor.close()

for i in list(frase4):
    print(i, end='')
    #O stdout só é atualizado quando há nova linha e como nós estamos mandando tudo na mesma é preciso forçar a atualização.
    sys.stdout.flush()
    time.sleep(0.05)