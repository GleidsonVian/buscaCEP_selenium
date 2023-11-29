from selenium import webdriver as opselenium
from selenium.webdriver.common.keys import Keys
import pyautogui as ptg
from selenium.webdriver.common.by import By

navegador = opselenium.Chrome()
navegador.get('https://buscacepinter.correios.com.br/app/endereco/index.php')

ptg.sleep(2)

navegador.find_element(By.NAME, 'endereco').send_keys('29156048')

ptg.sleep(2)

navegador.find_element(By.NAME, 'btn_pesquisar').click()

ptg.sleep(2)

rua = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]')[0].text
print(rua)

bairro = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[2]')[0].text
print(bairro)

cidade = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[3]')[0].text
print(cidade)

cep = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[4]')[0].text
print(cep)

# -----------------------------------------------------------

from openpyxl import load_workbook
import os

nome_arquivo_cep = 'C:\\Users\\Gleidson\\OneDrive - Sociedade Educacional do Espírito Santo - UVV\\Desktop\\PythonES\\Jupyter Notebooks\\arquivos excel\\extraindo dados de cep\\PesquisaEndereco_1.xlsx'
planilhaDadosEndereco = load_workbook(nome_arquivo_cep)

sheet_selecionada = planilhaDadosEndereco['Dados']

linha = len(sheet_selecionada['A']) +1
colunaA = 'A' + str(linha)
colunaB = 'B' + str(linha)
colunaC = 'C' + str(linha)
colunaD = 'D' + str(linha)

sheet_selecionada[colunaA] = rua
sheet_selecionada[colunaB] = bairro
sheet_selecionada[colunaC] = cidade
sheet_selecionada[colunaD] = cep

planilhaDadosEndereco.save(filename= nome_arquivo_cep)

os.startfile(nome_arquivo_cep)
