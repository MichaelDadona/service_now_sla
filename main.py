import function
from datetime import datetime
import schedule
import time


function.name()

def scheduler_inc():
    try:
        colunas = ["NÚMERO", "STATUS", "DATA DE INÍCIO", "TEMPO RESTANTE", "SLA"] # column name
        lista_url = function.fila_inicidentes()
        attr = 'INC' # search attribute for update_sla
        navegador = function.sso()
        function.update(navegador, lista_url, attr)
        print('\n', datetime.now().strftime("%d/%m/%y %H:%M:%S"))

        print('::::::: Executando FILA_1 ::::::::::::')
        function.service_now(navegador, lista_url[0], colunas)

        print('::::::: Executando FILA_2 ::::::::::::')
        function.service_now(navegador, lista_url[1], colunas)

        print('::::::: Executando FILA_3 ::::::::::::')
        function.service_now(navegador, lista_url[2], colunas)

        navegador.quit()
    except:
        pass
    
def scheduler_req():
    try:
        colunas = ["NÚMERO", "STATUS", "ABERTO EM", "TEMPO RESTANTE", "SLA"]
        lista_url = function.fila_requisicoes()
        attr = 'RIT'
        navegador = function.sso()
        function.update(navegador, lista_url, attr)
        print('\n', datetime.now().strftime("%d/%m/%y %H:%M:%S"))

        print('::::::: Executando FILA_1 ::::::::::::')
        function.service_now(navegador, lista_url[0], colunas)

        print('::::::: Executando FILA_2 ::::::::::::')
        function.service_now(navegador, lista_url[1], colunas)

        print('::::::: Executando FILA_3 ::::::::::::')
        function.service_now(navegador, lista_url[2], colunas)

        navegador.quit()
    except:
        pass

def scheduler_mud():
    try:
        colunas = ["NÚMERO", "INÍCIO", "TÉRMINO", "STATUS", "GRUPO"]
        lista_url = function.fila_mudancas()
        navegador = function.sso()
        print('\n', datetime.now().strftime("%d/%m/%y %H:%M:%S"))

        print('::::::: Executando FILA_1 ::::::::::::')
        function.service_now(navegador, lista_url[0], colunas)

        print('::::::: Executando FILA_2 ::::::::::::')
        function.service_now(navegador, lista_url[1], colunas)

        print('::::::: Executando FILA_3 ::::::::::::')
        function.service_now(navegador, lista_url[2], colunas)

        navegador.quit()
    except:
        pass



schedule.every(5).minutes.do(scheduler_inc)
schedule.every(60).minutes.do(scheduler_req)
schedule.every(60).minutes.do(scheduler_mud)


while 1:
    schedule.run_pending()
    time.sleep(1)
    #scheduler_inc()
    #scheduler_req()
    #scheduler_mud()