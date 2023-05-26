import pandas as pd 
import win32com.client as win32
from time import sleep
import json

def main():
    with open('conteudo.html','r', encoding="UTF-8") as arqConteudo:
        conteudo = arqConteudo.read()
    with open('conteudo.html','r', encoding="UTF-8") as arqAssunto:
        linhas = arqAssunto.readlines()
        for i in linhas:
            if ('<title>' in i):
                assunto = i.split('<title>')
                assunto = assunto[1].split('</title>')[0]
    with open('config.json', 'r', encoding="UTF-8") as arq:
        endereco = json.load(arq)['endereco']
    if (endereco == ""):
        endereco = input('entre com o endereço da planilha:')
        with open('config.json', 'w', encoding="UTF-8") as arq:
            i = {'endereco':f'{endereco}'}
            json.dump(i, arq)
    
    arqExcel = pd.read_excel(endereco, engine="openpyxl")
    cont = 0
    for receptor in arqExcel["Emails"]:
        try:
            enviaEmail(receptor, assunto, conteudo)
            cont += 1
        except:
            print(f'Erro ao enviar para {i}') 

    print(f'{cont} emails enviados')


def enviaEmail(receptor,assunto,conteudo):
    outlook = win32.Dispatch('outlook.application')

    email = outlook.CreateItem(0)

    email.To = receptor
    email.Subject = assunto
    email.HTMLBody = conteudo

    email.Send()

    print(f'Email enviado com sucesso para {receptor}')
    sleep(30)
    
main()
