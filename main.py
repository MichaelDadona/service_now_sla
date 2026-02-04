import pyautogui
import clipboard
import pyperclip
import time
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.utils import column_index_from_string

# === MAPA DE NOMES DOS MESES ===
nomes_meses = {
    '01': 'Janeiro',
    '02': 'Fevereiro',
    '03': 'Março',
    '04': 'Abril',
    '05': 'Maio',
    '06': 'Junho',
    '07': 'Julho',
    '08': 'Agosto',
    '09': 'Setembro',
    '10': 'Outubro',
    '11': 'Novembro',
    '12': 'Dezembro'
}

print('\n***************** Atenção *********************')
print('***   1 DEVE ESTAR NA TELA PRIMARY MENU       ***')
print('***   2 NÃO PODE ESTAR LOGADO NO WV - EOS     ***')
print('***   3 TERMINAL NEXUS DEVE ESTAR MAXIMIZADO  ***')
print('***   4 CYBERARK DEVE ESTAR EM MODO JANELA    ***')
print('***   NÃO USAR TECLADO E MOUSE DURANTE 60 SEG ***')
print('*************************************************')
print('***   CASO O PROGRAMA SEJA INTERROMPIDO, PODE ***')
print('***   FICAR TRAVADO O TECLADO, APERTE SHIFT   ***')
print('***   ESQUERDO + SHIFT DIREITO PARA DESTRAVAR ***')
print('*************************************************')

# === INPUT DO MÊS ===
while True:
    mes_input = input("Digite o número do mês (01 a 12): ").zfill(2)
    if mes_input in nomes_meses:
        break
    print("❌ Mês inválido. Tente novamente com valores de 01 a 12.")

# === CAMINHO EXCEL BASEADO NO MÊS ===
nome_mes = nomes_meses[mes_input]
caminho_excel = fr'\\rctr061\ito\PCP\APT\SCH MF\POSIC_SEC\VANS_POS\2026\01 - VANS CORTE 02 HORAS 01 - {mes_input} {nome_mes} 2026.xlsx'
print('***   NÃO USAR TECLADO E MOUSE DURANTE 60 SEG ***')

time.sleep(3)


######################### COLETAR DADOS ###############################

def enviar_comando(janela_titulo):
    janela = pyautogui.getWindowsWithTitle(janela_titulo)[0]
    janela.activate()
    time.sleep(1)
    janela.maximize()
    time.sleep(1)

    list_rot = ['VNB005T.', 'VNB062T.', 
                 'VNB038T.','VNB023T.',
                 'VNB058T.','VNB066T.',
                 'VNB008T1','VNB039T1','VNB055T1']
    list_valor = []

    pyautogui.press('home')
    time.sleep(1)
    pyautogui.press('end')
    time.sleep(1)
    pyautogui.write('WV')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.write('z')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.write('3')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    for rotina in list_rot:
        pyautogui.press('home')
        time.sleep(1)
        pyautogui.press('tab', presses=3)
        time.sleep(1)
        pyautogui.write(rotina)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.keyDown('shiftleft')
        pyautogui.keyDown('shiftright')
        pyautogui.press('tab')
        pyautogui.keyUp('shiftleft')
        pyautogui.keyUp('shiftright')
        pyautogui.press('up', presses=4)
        pyautogui.keyDown('shiftleft')
        pyautogui.keyDown('shiftright')
        pyautogui.press('right', presses=36)
        time.sleep(0.8)
        pyautogui.press('down', presses=4)
        pyautogui.keyUp('shiftleft')
        pyautogui.keyUp('shiftright')
        pyautogui.keyDown('ctrl')
        pyautogui.press('c')
        pyautogui.keyUp('ctrl')
        time.sleep(1)
        list_valor.append(pyperclip.paste())
        time.sleep(1)
        pyautogui.press('f3', presses=2)
        time.sleep(1)

    pyautogui.press('pageup')

    # Data de hoje no formato DDMMYY (para comparação)
    data_hoje = datetime.now().strftime('%d%m%y')


    caminho_arquivo = r'\\rctr061\ito\PCP\APT\SCH MF\POSIC_SEC\VANS_POS\vans.txt'
    with open(caminho_arquivo, 'w', encoding='utf-8') as f:
        for valor in list_valor:
            for linha in valor.splitlines():
                linha = linha.strip()
                if len(linha) > 28:
                    # Extrair a data que aparece após o ' W ' (ex: 060625)
                    partes = linha.split()
                    if len(partes) >= 5:
                        data_linha = partes[3]  # Supondo que a data está sempre na 4ª posição (índice 3)
                        if data_linha == data_hoje:
                            nome = linha[:9].strip()
                            horario_raw = linha[-6:].strip().zfill(6)
                            nova_linha = f"{nome} {horario_raw[:2]}:{horario_raw[2:]}"
                            nova_linha = nova_linha[:-2]  # Remove os dois últimos caracteres
                            f.write(nova_linha + '\n')
                            print(nova_linha)

janela_titulo = 'Conexão de Área de Trabalho Remota'
enviar_comando(janela_titulo)

######################### PREENCHER PLANILHA ###############################

caminho_txt = r'\\rctr061\ito\PCP\APT\SCH MF\POSIC_SEC\VANS_POS\vans.txt'

valores = {}
with open(caminho_txt, 'r', encoding='utf-8') as f:
    for linha in f:
        partes = linha.strip().split()
        if len(partes) == 2:
            chave, valor = partes
            valores[chave.strip()] = valor.strip()

hoje = datetime.today().date()

mapeamento = {
    2: [("VNB005T5", "B"), ("VNB005T1", "G"), ("VNB005T2", "L"), ("VNB005T3", "Q"), ("VNB005T4", "V")],
    3: [("VNB062T5", "B"), ("VNB062T1", "G"), ("VNB062T2", "L"), ("VNB062T3", "Q"), ("VNB062T4", "V")],
    4: [("VNB038T5", "B"), ("VNB038T1", "G"), ("VNB038T2", "L"), ("VNB038T3", "Q"), ("VNB038T4", "V")],
    5: [("VNB023T5", "B"), ("VNB023T1", "G"), ("VNB023T2", "L"), ("VNB023T3", "Q"), ("VNB023T4", "V")],
    6: [("VNB058T5", "B"), ("VNB058T1", "G"), ("VNB058T2", "L"), ("VNB058T3", "Q"), ("VNB058T4", "V")],
    7: [("VNB066T5", "B"), ("VNB066T1", "G"), ("VNB066T2", "L"), ("VNB066T3", "Q"), ("VNB066T4", "V")],
    8: [("VNB008T1", "B")],
    9: [("VNB039T1", "B")],
    10: [("VNB055T1", "B")],
    
}

wb = load_workbook(caminho_excel)

for aba_idx, rotinas in mapeamento.items():
    aba_nome = wb.sheetnames[aba_idx - 1]
    ws = wb[aba_nome]

    for rotina, col_letra in rotinas:
        col_idx = column_index_from_string(col_letra)
        col_data_idx = col_idx - 1

        for row in range(4, ws.max_row + 1):
            celula_data = ws.cell(row=row, column=col_data_idx)
            valor_data = celula_data.value

            data_celula = None
            if isinstance(valor_data, datetime):
                data_celula = valor_data.date()
            elif isinstance(valor_data, str):
                try:
                    data_celula = datetime.strptime(valor_data.strip(), "%d/%m/%Y").date()
                except ValueError:
                    try:
                        data_celula = datetime.strptime(valor_data.strip(), "%d/%b").replace(year=hoje.year).date()
                    except ValueError:
                        continue

            if data_celula == hoje:
                if rotina in valores:
                    ws.cell(row=row, column=col_idx).value = valores[rotina]

wb.save(caminho_excel)
print("✔ Planilha atualizada.")


#   VNB038T1     23794 W 080625 131442          366  0 0                        
#                                                    ABE ***  ABEND  =S522 U0000
#   VNB038T2     24974 W 080625 172450          366  0 0                        
#                                                    ABE ***  ABEND  =S522 U0000
#   VNB038T3     25808 W 080625 194453          366  0 0                        
#                                                    ABE ***  ABEND  =S522 U0000



# TSO RELCONN

# //*-------------------------------------------------------------------*
# //* STEP 1 - EXECUTA A REXX RELCONN 
# //*-------------------------------------------------------------------*
# //STEP1    EXEC PGM=IKJEFT01,PARM='TSO REXX RELCONN 01 H 01'
# //SYSEXEC  DD DISP=SHR,DSN=SYSX.PRD.USRCLIB                
# //SYSTSPRT DD SYSOUT=*
# //SYSTSIN  DD DUMMY

# //***********************************************                             
# //*        ENVIA E-MAILS                                                      
# //***********************************************                             
# //STEP2     EXEC PGM=IEBGENER                                                 
# //SYSPRINT  DD SYSOUT=*                                                       
# //OSMTP OUTPUT USERDATA=('TO=MICHAEL.DADONA@TIVIT.COM',                       
# //             'SUBJECT=VAN')                                        
# //SYSUT1    DD DSN=ISPP.PS431B.TEMP.REL.RELCONN,                                 
# //             DISP=SHR                                                       
# //SYSUT2    DD SYSOUT=A,DEST=OPTG0010,OUTPUT=*.OSMTP,                         
# //             OUTLIM=500000,DCB=(LRECL=200,RECFM=FB),FREE=CLOSE               
# //SYSIN     DD DUMMY                                                          
# //*                    




# PBRTS.BRTLB.REXX
# HM - SYSP.PRD.USRCLIB

# TSO RELCONN   %TVTFR119
# BRPAN5BD

# TSO ISRDDN


#    > PBRCV.BCCLB.VISA.JOBLIB     - JOBLIB
#    > PBRCV.BCCLB.VISA.PROCLIB    - PROCLIB

# SYSP.CTM.OVERLIB(DADAREXX)




#  pyinstaller --onefile Preencher_VANS.py
