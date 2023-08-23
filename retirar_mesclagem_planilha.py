import os
import win32com.client as win32

def retirar_mesclagem(url_planilha):
    xl = win32.Dispatch("Excel.Application")

    # Obtendo o path atual do arquivo
    path_atual = os.getcwd()
    planilha_com_VBA = "retirar_mesclagem_planilhas.xlsm"

    # Abrindo o arquivo
    if os.path.isfile(os.path.join(path_atual, planilha_com_VBA)):
        wb = xl.Workbooks.Open(os.path.join(path_atual, planilha_com_VBA))
    else:
        print("Arquivo não encontrado.")
        return

    # Executando o código VBA diretamente no Excel
    try:
        xl.Run("DesmesclarCelulas", url_planilha)
    except: 
        print("Erro ao executar a função VBA 'retirar_mesclagem'.")

    # Salvando e fechando o arquivo
    wb.Save()
    wb.Close()

    # Fechando o Excel
    xl.Quit()