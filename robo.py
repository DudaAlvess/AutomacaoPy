from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import openpyxl
from time import sleep
import pandas
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd 

driver = webdriver.Chrome()

driver.get("http://syscassol.cassol.local:8180/cas-syscassol/faces/login.xhtml")

driver.maximize_window()

email_input = driver.find_element("id", "frmLogin\:loginNome")
email_input.send_keys("maria.alves")
password_input = driver.find_element("id", "frmLogin\:loginSenha")
password_input.send_keys("M@ria2006")
password_input.send_keys(Keys.RETURN)
sleep(1)

driver.get("http://syscassol.cassol.local:8180/cas-syscassol/faces/pages/logistica/CalculoFrete/CrudCasFreteDesconto.xhtml")

df = pd.read_excel(r"C:\Users\maria.alves\Desktop\robo.xlsx", sheet_name='Plan2')

for i, row in df.iterrows():
    pedido = row['Pedido']
    frota = row['Frota']
    carga = row['Carga ']
    valor = row['Valor Desconto']
    observacao = row['Observação']



    vl_tratado_adicional = str(round(valor,2))
    vl_tratado_adicional = vl_tratado_adicional.replace(".",",")
    print(pedido)

    input_pedido = driver.find_element('xpath','//*[@id="form:numeroPedido_input"]')
    input_pedido.clear()
    input_pedido.send_keys(pedido)

    botao_pesquisa = driver.find_element('xpath','//*[@id="form:j_idt48"]/span')
    botao_pesquisa.click()

    botao_adicionar = driver.find_element('xpath','//*[@id="form:j_idt50"]/span')
    botao_adicionar.click()

    sleep(2)
    print('Frota')
    local_frota = driver.find_element('xpath','//*[@id="form:salvarFrota_input"]')
    local_frota.send_keys(frota)

    click_conf = driver.find_element('xpath','//*[@id="form:quinzenaDoItem"]')
    click_conf.click()
    
    sleep(1)
    print('Quinzena')
    click_conf = driver.find_element('xpath','//*[@id="form:quinzenaDoItem"]')
    click_conf.click()

    click_quinz = driver.find_element('xpath','//*[@id="form:quinzenaDoItem_61"]')
    click_quinz.click()

    click_conf = driver.find_element('xpath','//*[@id="form:salvarCarga_input"]')
    click_conf.click()

    print('indo carga')
    local_carga = driver.find_element('xpath','//*[@id="form:salvarCarga_input"]')
    local_carga.send_keys(carga)
    
    print('pedido')
    local_pedido = driver.find_element('xpath','//*[@id="form:salvarPedido_input"]')
    local_pedido.send_keys(pedido)

    print('valor')
    local_valor = driver.find_element('xpath','//*[@id="form:salvarValor_input"]')
    local_valor.send_keys(vl_tratado_adicional)

    print('observação')
    local_obs = driver.find_element('xpath','//*[@id="form:salvarObservacao"]')
    local_obs.send_keys(observacao)

    print('gravando...')
    click_gravar = driver.find_element('xpath','//*[@id="form:j_idt67"]')
    click_gravar.click()
    
    print('Lançado!')
  
    sleep(2)
    
  