import os
import win32com.client
import pyautogui
from datetime import datetime


pasta_entrada = r'C:\Python\Pasta_entradaDOC'
pasta_saida = r'C:\Python\pasta_saidapdf'

word = win32com.client.Dispatch('Word.Application')
word.Visible = False  # ou True se quiser ver

for nome_arquivo in os.listdir(pasta_entrada):
    if nome_arquivo.lower().endswith(('.doc', '.docx')):
        caminho_arquivo = os.path.join(pasta_entrada, nome_arquivo)
        nome_pdf = os.path.splitext(nome_arquivo)[0] + '.pdf'
        caminho_pdf = os.path.join(pasta_saida, nome_pdf)

        try:
            doc = word.Documents.Open(caminho_arquivo)
            doc.SaveAs(caminho_pdf, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close()
            print(f'Convertido: {nome_arquivo}')
        except Exception as e:
            print(f'Erro ao converter: {nome_arquivo} + {e}')

word.Quit()

#Ultilizando o Pyautogui
pyautogui.PAUSE = 2
pyautogui.press('win')
pyautogui.write('Explorador de Arquivos')
pyautogui.press('enter')
pyautogui.click(x=114, y=795)
pyautogui.click(x=278, y=486, clicks=2)
pyautogui.click(x=423, y=258, clicks=2)

#Data da conversao
data_hoje = datetime.now()
data_formatada = data_hoje.strftime('%d/%m/%Y')
print(data_formatada)
print("Conversão Concluída")

print("Mais um projeto de PEDRO EZEQUIEL")
