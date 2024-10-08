import pandas as pd
import glob
import os
import shutil
import sqlalchemy.types
import win32com.client
import time
from datetime import datetime
from sqlalchemy import create_engine
from selenium import webdriver
from selenium.webdriver.common.by import By

# CAMIHO DO ARQUIVO A SER BAIXADO PRIMARIAMENTE
caminho = r'E:\Automação'

# ESCOLHENDO AS PREFERÊNCIAS DO DRIVER
driver_path = r'E:\python user\msedgedriver.exe'
prefs = {
    "download.default_directory": caminho,
    "download.prompt_for_download": False,
    "safebrowsing.enables": True
}
options = webdriver.EdgeOptions()
options.add_experimental_option("prefs", prefs)
driver = webdriver.Edge(options=options)

# ACESSANDO O SITE E CLICANDO NO BAIXADOR DO ARQUIVO
driver.get("http://meusite.salvador.ba.gov.br/admin/index.php")
time.sleep(2)
driver.find_element(By.NAME, "email").send_keys("usuario")
driver.find_element(By.NAME, "senha").send_keys("senha")
driver.find_element(By.XPATH, "/html/body/div/div/form/div/div[6]/center/button").submit()
time.sleep(5)

driver.get("http://meusite.salvador.ba.gov.br/admin/sistema/app/lista.php")
time.sleep(3)
driver.find_element(By.ID, "btn_arquivo_excel").click()
print("Baixando arquivo...")
time.sleep(1300)
print("Arquivo Baixado!")


# ENCONTRANDO O AQRUIVO DE EXTENSÃO XLS MAIS RECENTE
arquivo_xls = max(glob.glob(os.path.join(caminho, '*.xls')), key=os.path.getctime)

if not arquivo_xls:
    print('O arquivo não existe na pasta')

else:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Open(arquivo_xls)

    # Essa linha desativa o Verificador de Compatibilidade
    excel.DisplayAlerts = False

    arquivo_xls_compatible = os.path.join(caminho, 'arquivooriginal_compatible.xls')
    workbook.SaveAs(arquivo_xls_compatible, FileFormat=56)

    workbook.Close()
    excel.Quit()
    print('Compatibilidade concluída')

    # ATUALIZAR TABLE planilhaGRE COM A PLANILHA GRE 2.0
    conn_str = (
            "mssql+pyodbc://"
            + "user_name:senha"
            + "ip_banco/nome_banco"
            + "?driver=ODBC+Driver+17+for+SQL+Server"
    )
    engine = create_engine(conn_str)

    caminho_arquivo_excel = r'E:\Automação\arquivooriginal_compatible.xls'
    df = pd.read_excel(caminho_arquivo_excel)

    # Nome de todas as colunas (e sua ordem) da table
    colunas_tabela = [
        "n_mapa",
        "remanejada",
        "etapa",
        "vale",
        "trecho",
        "selagem",
        "grupo_familiar",
        "prioridade",
        "processo",
        "reuniao"
    ]

    df.columns = colunas_tabela
    dtype = {column: sqlalchemy.types.Text() for column in df.columns}
    df.to_sql(name='planilhaGRE', con=engine, if_exists='replace', index=False, dtype=dtype)
    engine.dispose()

    print('Atualização da table planilhaGRE realizada')
    time.sleep(2)

    # ENVIAR ARQUIVO CONVERTIDO PARA PASTA DE ATUALIZAÇÕES
    pasta_att = r'E:\Automação\att planilhaGRE'
    data_hora_agora = datetime.now().strftime("%d-%m-%Y-%H%M-%S")

    # RENOMEAR O ARQUIVO
    arquivo_compatible = r'E:\Automação\arquivooriginal_compatible.xls'
    novo_arquivo_compatible = f'planilhaGRE_{data_hora_agora}.xls'
    shutil.move(arquivo_compatible, os.path.join(pasta_att, novo_arquivo_compatible))
    print('Arquivo convertido enviado para pasta de atualizações')
    time.sleep(2)

# REMOVER O ARQUIVO XLS PRINCIPAL
    os.remove(arquivo_xls)
    print('Arquivo primário XLS excluído')
    time.sleep(3)

# ENVIAR POR EMAIL ESSE ARQUIVO COMPATÍVEL
    arquivo_para_envio = max(glob.glob(os.path.join(r'E:\Automação\att planilhaGRE', '*.xls')), key=os.path.getctime)

    driver.get("https://accounts.google.com/v3/signin/")
    driver.find_element(By.ID, "identifierId").send_keys("email")
    driver.find_element(By.ID, "identifierNext").click()
    time.sleep(10)
    driver.find_element(By.NAME, "usuário").send_keys("senha")
    driver.find_element(By.ID, "passwordNext").click()
    time.sleep(10)

    print('Gmail acessado')
    driver.get("https://mail.google.com/mail/u/0/#inbox?compose=new")
    time.sleep(10)
    driver.find_element(By.CLASS_NAME, "aFw").send_keys("email1@gmail.com;email2@gmail.com;email3@gmail.com")
    driver.find_element(By.NAME, "subjectbox").send_keys("Planilha de atualização 2.0 COMPLETA (Não responder)")
    driver.find_element(By.CLASS_NAME, "editable").send_keys(
        f"Olá, \n Segue em anexo planilha de atualização 2.0 completa \n\n\n\n   ~ Mensagem Automática ~ \n\n"
        f"*** Deletar esta mensagem após o download da planilha para não ocupar espaço na caixa de email *** ")
    time.sleep(5)
    driver.find_element(By.XPATH, '//input[@type="file"]').send_keys(fr"{arquivo_para_envio}")
    time.sleep(10)
    driver.find_element(By.CLASS_NAME, "aoO").click()
    print('carregado_e_enviado')
    time.sleep(3)
    driver.quit()
    print('Programa finalizado')
