import openpyxl
import os

def salvar_excel(data_hora, temperatura, umidade):
    nome_arquivo = 'planilha_excel.xlsx'

    if not os.path.exists(nome_arquivo):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Dados'
        ws.append(['Data e Hora', 'Temperatura', 'Umidade'])
    else:
        wb = openpyxl.load_workbook(nome_arquivo)
        ws = wb.active

    ws.append([data_hora, temperatura, umidade])
    wb.save(nome_arquivo)

