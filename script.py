import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time

# Configura o chromedriver para executar o serviço
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

navegador.get("http://192.168.0.65/Biblivre5/?action=search_bibliographic")

# Login
usuario = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menu"]/ul/li[5]/input[1]')))
usuario.send_keys("admin")

senha = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menu"]/ul/li[5]/input[2]')))
senha.send_keys("abracadabra")

btnEntrar = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menu"]/ul/li[4]/button')))
btnEntrar.click()

time.sleep(1)
###############################################################################################################################################################

# Pegar dados de uma planilha
df = pd.read_excel('certa Tombo Livros - ETEC Emilio Hernandes Aguilare).xls')
print(df)
for indice, linha in df.iterrows():
    navegador.get("http://192.168.0.65/Biblivre5/?action=cataloging_bibliographic")
    time.sleep(1)
    btnNovoRegistro = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="new_record_button"]')))
    btnNovoRegistro.click()
    time.sleep(1)
    def inserir_dado(campo_xpath, valor, default=""):
        if pd.isna(valor) or valor == "":
            valor = default
        campo = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, campo_xpath)))
        campo.clear()
        campo.send_keys(str(valor))
    
    inserir_dado('//*[@id="biblivre_form"]/div[13]/fieldset/div[2]/div[2]/div[2]/input', linha['Autor (es)'])
    inserir_dado('//*[@id="biblivre_form"]/div[20]/fieldset/div[2]/div[5]/div[2]/input', linha['Responsável (is) / Colaboadores'])
    inserir_dado('//*[@id="biblivre_form"]/div[20]/fieldset/div[2]/div[3]/div[2]/input', linha['Título'])
    inserir_dado('//*[@id="biblivre_form"]/div[36]/fieldset/div[2]/div[2]/div[2]/input', linha['Série - Coleção'])
    inserir_dado('//*[@id="biblivre_form"]/div[27]/fieldset/div[2]/div[1]/div[2]/input', linha['Local de publicação'])
    inserir_dado('//*[@id="biblivre_form"]/div[27]/fieldset/div[2]/div[2]/div[2]/input', linha['Editora'])
    inserir_dado('//*[@id="biblivre_form"]/div[27]/fieldset/div[2]/div[3]/div[2]/input', linha['Ano de Publicação'])
    inserir_dado('//*[@id="biblivre_form"]/div[22]/fieldset/div[2]/div[1]/div[2]/input', linha['Edição'])
    inserir_dado('//*[@id="biblivre_form"]/div[5]/fieldset/div[2]/div[2]/div[2]/input', linha['Língua'])
    inserir_dado('//*[@id="biblivre_form"]/div[2]/fieldset/div[2]/div/div[2]/input', linha['ISBN'], default="Não informado")
    inserir_dado('//*[@id="biblivre_form"]/div[10]/fieldset/div[2]/div[1]/div[2]/input', linha['Classificação'])
    inserir_dado('//*[@id="biblivre_form"]/div[10]/fieldset/div[2]/div[2]/div[2]/input', linha['Cutter'])
    inserir_dado('//*[@id="biblivre_form"]/div[57]/fieldset/div[2]/div[1]/div[2]/input', linha['Assunto 1'])

    if pd.notna(linha['Assunto 2']):
        novo_campo = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="biblivre_form"]/div[57]/fieldset/legend/a[2]')))
        novo_campo.click()
        inserir_dado('//*[@id="biblivre_form"]/div[57]/fieldset[2]/div[2]/div[1]/div[2]/input', linha['Assunto 2'])

    if pd.notna(linha['Assunto 3']):
        novo_campo = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="biblivre_form"]/div[57]/fieldset[1]/legend/a[2]')))
        novo_campo.click()
        inserir_dado('//*[@id="biblivre_form"]/div[57]/fieldset[2]/div[2]/div[1]/div[2]/input', linha['Assunto 3'])

    if pd.notna(linha['Exemplar']):
        try:
            exemplar_valor = int(linha['Exemplar'])
            inserir_dado('//*[@id="cataloging_search"]/div[5]/div[2]/div[4]/div/fieldset/div[1]/div[2]/input', exemplar_valor)
        except ValueError:
            print(f"Valor inválido para Exemplar na linha {indice}: {linha['Exemplar']}")

    time.sleep(1)

    btnSalvar = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="cataloging_search"]/div[4]/div/div[1]/div[3]/a[1]')))
    btnSalvar.click()

    time.sleep(1)

    if linha['n de Tombo'] == 3494:
        navegador.exit(0)

# Pro navegador não fechar
while True:
    pass
