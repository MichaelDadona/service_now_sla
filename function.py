import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import re
import time
import pyperclip
from config import CHROME_PROFILE_PATH


#::::::::::::::::: Filas Incidentes ::::::::::::::::
fila_inc_1 = ('LINK_FILAS', 'LINK_TEAMS')
fila_inc_2 = ('LINK_FILAS', 'LINK_TEAMS')
fila_inc_3 = ('LINK_FILAS', 'LINK_TEAMS')

#::::::::::::::::: Filas Requisições ::::::::::::::
fila_req_1 = ('LINK_FILAS', 'LINK_TEAMS')
fila_req_2 = ('LINK_FILAS', 'LINK_TEAMS')
fila_req_3 = ('LINK_FILAS', 'LINK_TEAMS')

#::::::::::::::::: Filas Mudanças :::::::::::::::::
fila_mud_1 = ('LINK_FILAS', 'LINK_TEAMS')
fila_mud_2 = ('LINK_FILAS', 'LINK_TEAMS')
fila_mud_3 = ('LINK_FILAS', 'LINK_TEAMS')

def fila_inicidentes():
    lista_url = [fila_inc_1, fila_inc_2, fila_inc_3]
    return lista_url

def fila_requisicoes():
    lista_url = [fila_req_1, fila_req_2, fila_inc_3]
    return lista_url

def fila_mudancas():
    lista_url = [fila_mud_1, fila_mud_2, fila_mud_3]
    return lista_url




def name():
    print(''' █▀▀▀█ █░░░ █▀▀█
▀▀▀▄▄ █░░░ █▄▄█
█▄▄▄█ █▄▄█ █░▒█
_______________''')

# convert list elements to text
def conv_text(lista, elemento):
    return [lista.append(_.text) for _ in elemento]

# correct errors in the text
def substitui(newlist, lista):
    return [newlist.append(_.replace('\n', '')) for _ in lista]

# divided the list into sublists with the number of elements informed
def div_lista(lista, tamanho):
    return [lista[_:_+tamanho] for _ in range(0,len(lista),tamanho)]

# convert data into table
def conv_table(dados, colunas):
    tabela = pd.DataFrame(dados)
    tabela = tabela.drop(5, axis=1)
    tabela = tabela.rename(columns={0: colunas[0], 1: colunas[1], 2: colunas[2], 3: colunas[3], 4: colunas[4]})
    tabela.index += 1
    tabela = pd.DataFrame.to_string(tabela)
    return tabela

# log into Service Now and set browser parameters
def sso():
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('window-size=700,657')
        options.add_argument('window-position=0,0')
        options.add_argument(CHROME_PROFILE_PATH)
        options.add_experimental_option('excludeSwitches', ['enable-logging']) # remove error messages
        var = os.path.dirname(os.path.realpath('__file__'))  # get the file path
        navegador = webdriver.Chrome(
            executable_path=var + r'\chromedriver.exe', options=options)

        sso = 'LINK_SSO'        
        navegador.implicitly_wait(2)
        navegador.get(sso)
        navegador.find_element('xpath', '//*[@id="loginPage"]/div/div/button').click()
        time.sleep(2)
        return navegador
    except:
        print('ERRO SSO')
        pass

# enters all occurrences to update the SLA value        
def update(navegador, lista_url, attr):
    try:
        navegador.implicitly_wait(3)
        page_update = 'LINK_PAGE_UPDATE_SLA'
        lista_elementos = []
        for _ in lista_url:
            navegador.get(_[0])
            x = navegador.find_elements(By.PARTIAL_LINK_TEXT, attr)
            conv_text(lista_elementos, x)

        if lista_elementos:
            navegador.get(page_update)
            navegador.find_element('xpath',
                                   '/html/body/div[5]/div/div/header/div[1]/div/div[2]/'
                                   'div/div[4]').click()
            for _ in lista_elementos:
                navegador.find_element('xpath', '//*[@id="sysparm_search"]').send_keys(Keys.LEFT_CONTROL, 'a')
                time.sleep(0.5)
                navegador.find_element('xpath', '//*[@id="sysparm_search"]').send_keys(_, Keys.ENTER)
                time.sleep(1.5)
        else:
            pass
    except:
        print('ERRO NO UPDATE_SLA')
        pass

# collects and treats table values        
def service_now(navegador, grupo, colunas):
    try:
        navegador.implicitly_wait(3)
        navegador.get(grupo[0])
        elementos = navegador.find_elements(By.CLASS_NAME, 'vt')
        if elementos:
            lista_txt = []
            conv_text(lista_txt, elementos)
            trata_txt = []
            substitui(trata_txt, lista_txt)
            sublistas = div_lista(trata_txt, 6)
            tabela = conv_table(sublistas, colunas)
            teams(navegador, grupo, tabela)
        else:
            pass
    except:
        print('ERROR SERVICE_NOW')
        pass

# delivers the cases to the Microsoft Teams group
def teams(navegador, grupo, tabela):
    try:
        navegador.implicitly_wait(30)
        navegador.get(grupo[1])
        print(tabela)
        _ = re.compile("[ ][7-9][0-9][,][0-9]?[0-9]?[%]") # occurrences between 70% to 99% will be sent as !!important!!
        alerta = re.search(_, tabela)
        navegador.switch_to.frame(navegador.find_element(By.XPATH,
                                                         "//iframe[contains(@class, 'embedded-page-content')]"))
        textbox = navegador.find_element(By.XPATH,
                                         "//div[@role='textbox' and contains(@class, 'cke_textarea_inline') "
                                         "and contains(@class, 'cke_editable')]")
        sendbutton = navegador.find_element(By.XPATH,
                                            "//button[@data-track-module-name='sendButton']")
        alertbutton = navegador.find_element(By.XPATH,
                                             '//*[@id="message-pane-layout-a11y"]/div[2]/div/div[4]/div[2]/div[1]/div[1]/button[2]')
        pyperclip.copy(tabela)
        time.sleep(1)

        if alerta:
            alertbutton.click()
            time.sleep(1)
            textbox.send_keys(Keys.CONTROL, 'v')
            time.sleep(1)
            sendbutton.click()
            time.sleep(2)
        else:
            textbox.send_keys(Keys.CONTROL, 'v')
            time.sleep(1)
            sendbutton.click()
            time.sleep(2)
    except:
        print('ERROR TEAMS')
        pass


