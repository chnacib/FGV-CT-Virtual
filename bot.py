from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import Tk, filedialog
from pathlib import Path
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
from time import sleep
from selenium import webdriver
from tkcalendar import Calendar
from PIL import ImageTk, Image
from datetime import date
import os
from tkinter.filedialog import askopenfilename
import urllib
import random



data_atual = date.today()
dataFormatada = data_atual.strftime('%d/%m/%Y')
day_today = int(data_atual.strftime('%d'))
month_today = int(data_atual.strftime('%m'))
year_today = int(data_atual.strftime('%y'))

listacpf = []
listanome = []
listadata = []
listahora = []
listadownload = []
listahardware = []
listaram = []
listainsta = []
listaverifica = []
listacelular = []
listamensagem = []
listamsgcheck = []
# Caminho para a raiz do projeto
ROOT_FOLDER = Path(__file__).parent.parent
# Caminho para a pasta onde o chromedriver está
CHROME_DRIVER_PATH = 'x'

class App:
    def make_chrome_browser(self,*options: str) -> webdriver.Chrome:
        chrome_options = webdriver.ChromeOptions()

        # chrome_options.add_argument('--headless')
        if options is not None:
            for option in options:
                chrome_options.add_argument(option)

        chrome_service = Service(
            executable_path=CHROME_DRIVER_PATH,
        )

        browser = webdriver.Chrome(
            service=chrome_service,
            options=chrome_options

        )
        a.browser = browser
        return browser
    def start_browser(self):
        browser_confirm = tk.Label(root, text='Abrindo browser', fg='red')
        browser_confirm.place(x=240,y=860,width=200,height=25)
        root.update()
        if __name__ == '__main__':
            a.options = ('--disable-gpu', '--no-sandbox',)
            a.browser = a.make_chrome_browser(*a.options)
            a.browser.get('x')


    def login_prepona(self):
        logprep_confirm = tk.Label(root, text='Efetuando login em Prepona', fg='red')
        logprep_confirm.place(x=280,y=860,width=200,height=25)
        root.update()
        inputlog = a.browser.find_element(By.NAME, 'Login')
        inputpasswd = a.browser.find_element(By.NAME, 'Senha')
        inputlog.send_keys('x')
        inputpasswd.send_keys('x')
        inputpasswd.send_keys(Keys.ENTER)
        sleep(2)

    def verificar_dispositivo(self):
        veridisp_confirm = tk.Label(root, text='Carregando candidatos', fg='red')
        veridisp_confirm.place(x=280,y=860,width=200,height=25)
        root.update()
        a.browser.get('x')
        sleep(3)

    def selecionar_filtros(self):
        filtro_confirm = tk.Label(root, text='Selecionando filtros', fg='red')
        filtro_confirm.place(x=280,y=860,width=200,height=25)
        root.update()
        input_date = a.browser.find_element(By.XPATH, '//*[@id="DataProgramada"]').clear()
        input_date = a.browser.find_element(By.XPATH, '//*[@id="DataProgramada"]').send_keys(a.cal_value)
        sleep(1)
        filter_btn = a.browser.find_element(By.XPATH, '//*[@id="btnVerificarDispositivoRemoto"]').click()
        sleep(2)
        filter2_btn = a.browser.find_element(By.XPATH,"//*[@href='/TCN/Monitoramento/ListaConsultarDispositivoRemoto?GridDispositivoRemoto-sort=IdSituacaoDispositivo-asc']").click()
        sleep(1)

    def verificar_candidatos(self):
        pb = ttk.Progressbar(
            root,
            orient='horizontal',
            mode='indeterminate',
            length=100
        )
        pb.place(x=280,y=860,width=200,height=25)
        pb.start()
        root.update()
        pos = 1
        for i in range(20):
            pb['value'] += 50
            root.update()
            try:
                candidato_select = a.browser.find_element(By.XPATH,
                                                        f'/html/body/div[1]/div/section[1]/div[7]/table/tbody/tr[{pos}]')
                candidato_select.click()
                sleep(1)
            except:
                break
            cpf_candidato = a.browser.find_element(By.XPATH,
                                                 f'//*[@id="GridVerificarDispositivoRemoto"]/table/tbody/tr[{pos}]/td[1]').text
            nome_candidato = a.browser.find_element(By.XPATH,
                                                  f'//*[@id="GridVerificarDispositivoRemoto"]/table/tbody/tr[{pos}]/td[2]').text
            data_agendamento = a.browser.find_element(By.XPATH,
                                                    f'//*[@id="GridVerificarDispositivoRemoto"]/table/tbody/tr[{pos}]/td[3]').text
            hora_agendamento = a.browser.find_element(By.XPATH,
                                                    f'//*[@id="GridVerificarDispositivoRemoto"]/table/tbody/tr[{pos}]/td[4]').text
            try:
                data_download = a.browser.find_element(By.XPATH,
                                                     f'//*[@id="GridVerificarDispositivoRemoto"]/table/tbody/tr[{pos}]/td[5]').text
                if not data_download:
                    data_download = "Não fez o download da prova"
            except:
                data_download = "Não fez o download da prova"
            sleep(1)
            try:
                processador_candidato = a.browser.find_element(By.XPATH,
                                                             f'//*[@id="GridDispositivoRemoto"]/table/tbody/tr[1]/td[3]').text
                ram_candidato = a.browser.find_element(By.XPATH,
                                                     f'//*[@id="GridDispositivoRemoto"]/table/tbody/tr[1]/td[5]').text
                if not processador_candidato:
                    processador_candidato = 'Não instalou o aplicativo'
                    ram_candidato = 'Não instalou o aplicativo'
            except:
                processador_candidato = 'Não instalou o aplicativo'
                ram_candidato = 'Não instalou o aplicativo'
            try:
                instal_data = a.browser.find_element(By.XPATH,
                                                   f'//*[@id="GridDispositivoRemoto"]/table/tbody/tr[1]/td[6]').text
                if not instal_data:
                    instal_data = 'Não instalou o aplicativo'
            except:
                instal_data = 'Não instalou o aplicativo'
            try:
                verifica_data = a.browser.find_element(By.XPATH,
                                                     f'//*[@id="GridDispositivoRemoto"]/table/tbody/tr[1]/td[7]').text
                if instal_data == 'Não instalou o aplicativo':
                    verifica_data = 'Não instalou o aplicativo'
                elif not verifica_data:
                    verifica_data = 'Não verificou o dispositivo'
            except:
                verifica_data = 'Não verificou o dispostivo'
            listacpf.append(cpf_candidato)
            listanome.append(nome_candidato)
            listadata.append(data_agendamento)
            listahora.append(hora_agendamento)
            listadownload.append(data_download)
            listahardware.append(processador_candidato)
            listaram.append(ram_candidato)
            listainsta.append(instal_data)
            listaverifica.append(verifica_data)

            pos = pos + 1
        pb.stop()
    def pular_pagina(self):
        try:
            next_page = a.browser.find_element(By.XPATH,
                                             '//*[@id="GridVerificarDispositivoRemoto"]/div/a[3]/span').click()
            sleep(2)
            a.verificar_candidatos()
        except:
            pass
    def open_cert(self):
        opencert_confirm = tk.Label(root, text='Abrindo CertPessoas', fg='red')
        opencert_confirm.place(x=280,y=860,width=200,height=25)
        root.update()
        a.browser.get('x')
        sleep(2)

    def cert_login(self):
        logcert_confirm = tk.Label(root, text='Efetuando login em CertPessoas', fg='red')
        logcert_confirm.place(x=280,y=860,width=200,height=25)
        root.update()
        inputlogin = a.browser.find_element(By.NAME, 'Login1$UserName')
        inputlogin.send_keys('x')
        inputpasswd = a.browser.find_element(By.NAME, 'Login1$Password')
        inputpasswd.send_keys('x')
        sleep(2)
        inputlogin.send_keys(Keys.ENTER)
    def get_celular(self):
        progress = ttk.Progressbar(root, orient=HORIZONTAL, length=100, mode='determinate')
        progress.place(x=280,y=860,width=200,height=25)
        start_value = 100/len(listacpf)
        progress_value = start_value
        for cpf in listacpf:
            progress['value'] = progress_value
            a.browser.get('xxxxxxxx')
            inputcpf = a.browser.find_element(By.XPATH, '//*[@id="ctl00_ContentMain_CPF"]')
            inputcpf.send_keys(cpf)
            searchbtn = a.browser.find_element(By.XPATH, '//*[@id="ctl00_ContentMain_btnProcurar"]').click()
            sleep(5)
            try:
                ddd_num = a.browser.find_element(By.XPATH,'//*[@id="ctl00_ContentMain_CandidatoCadastro1_DDDCelular"]').get_attribute('value')
                celular_num = a.browser.find_element(By.XPATH,'//*[@id="ctl00_ContentMain_CandidatoCadastro1_TelefoneCelular"]').get_attribute('value')
                celular = ddd_num + celular_num
                listacelular.append(celular)
            except:
                sleep(5)
                ddd_num = a.browser.find_element(By.XPATH,'//*[@id="ctl00_ContentMain_CandidatoCadastro1_DDDCelular"]').get_attribute('value')
                celular_num = a.browser.find_element(By.XPATH,'//*[@id="ctl00_ContentMain_CandidatoCadastro1_TelefoneCelular"]').get_attribute('value')
                celular = ddd_num + celular_num
                listacelular.append(celular)
            progress_value = progress_value + start_value
            root.update()
        concluido_msg = tk.Label(root,text='Process finished!',fg='green')
        concluido_msg.place(x=280,y=860,width=200,height=25)
        root.update()
    def data_get(self):
        cal_value = cal.get_date()
        cal_confirm = tk.Label(root,text='Saved!',fg='green')
        cal_confirm.place(x=370,y=470,width=40,height=20)
        a.cal_value = cal_value
    def importar_planilha(self):
        a = askopenfilename()  # Isto te permite selecionar um arquivo
        xls = pd.ExcelFile(a)
        df = pd.read_excel(a,sheet_name='Planilha1')
        a_path = os.path.abspath(a)
        planilha_path = tk.Label(root, text=a_path,fg='green')
        planilha_path.place(x=230,y=630,width=300,height=30)
    def criar_planilha(self):
        dict = {'CPF':listacpf,'Nome':listanome,'Data':listadata,'Hora':listahora,'Download':listadownload,'Hardware':listahardware,'RAM':listaram,'Instalação':listainsta,'Verificação':listaverifica,'Celular':listacelular,'Check':listamsgcheck}
        df_new = pd.DataFrame(dict)
        df_new.to_excel('ResultadoPrepona.xlsx')

    def check_if_open(self):
        checkopen_confirm = tk.Label(root,text='Aguardando QRCODE',fg='green')
        checkopen_confirm.place(x=280,y=860,width=200,height=25)
        root.update()
        try:
            check_wpp = a.browser.find_element(By.XPATH, '//*[@id="side"]')
        except:
            sleep(3)
            a.check_if_open()

    def check_situation(self):
        for nome, download, hardware, ram, install, verifica in zip(listanome, listadownload, listahardware, listaram,
                                                                    listainsta, listaverifica):
            if not ram == 'Não instalou o aplicativo':
                ram = int(ram)
            if install == 'Não instalou o aplicativo':
                mensagem = f"""Prezado(a), {nome}. 

    Mensagem A

    Obrigado!"""
                msg_check = "Não instalou o aplicativo"
            elif ram < 3800:
                mensagem = f"""Prezado(a), {nome}. 

    Mensagem B

    """
                msg_check = "RAM inferior"
            elif "Celeron" in str(hardware):
                mensagem = f"""Prezado(a), {nome}. 

    Mensagem C
    """
                msg_check = "Processador inferior"

            elif "Pentium" in str(hardware):
                mensagem = f"""Prezado(a), {nome}. 

    Mensagem D 
            """
                msg_check = "Processador inferior"

            elif "Dual" in str(hardware):
                mensagem = f"""Prezado(a), {nome}. 

    Mensagem E
            """
                msg_check = "Processador inferior"
            elif "Quad" in str(hardware):
                mensagem = f"""Prezado(a), {nome}. 

    Mensagem F
            """
                msg_check = "Processador inferior"

            elif verifica == 'Não verificou o dispositivo' and download == 'Não fez o download da prova':
                mensagem = f"""Prezado(a), {nome}. 
                
                
    Mensagem G
            """
                msg_check = "Não verificou o dispositivo e não fez o download da prova"
            elif verifica == 'Não verificou o dispositivo':
                mensagem = f"""Prezado(a), {nome}.


    Mensagem H
                """
                msg_check = "Não verificou o dispositivo"
            elif download == 'Não fez o download da prova':
                mensagem = f"""Prezado(a), {nome}. 

    Mensagem I
            
            """
                msg_check = "Não fez o download da prova"
            else:
                mensagem = "Não enviar"
                msg_check = "Não enviar"
            listamensagem.append(mensagem)
            listamsgcheck.append(msg_check)

    def send_msg(self):
        send_click = a.browser.find_element(By.XPATH,
                                          '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]')
        send_click.send_keys(Keys.ENTER)

    def send_wpp(self):
        progress = ttk.Progressbar(root, orient=HORIZONTAL, length=100, mode='determinate')
        progress.place(x=280, y=860, width=200, height=25)
        start_value = 100 / len(listacpf)
        progress_value = start_value
        for mensagem, celular in zip(listamensagem, listacelular):
            progress['value'] = progress_value
            mensagem = str(mensagem)
            celular = str(celular)
            mensagem_url = urllib.parse.quote(mensagem)
            if mensagem != "Não enviar" and len(celular) >= 11:
                try:
                    link = f"https://web.whatsapp.com/send?phone=55{celular}&text={mensagem_url}"
                    a.browser.get(link)
                    sleep(random.randint(13, 15))
                    try:
                        checkvalid = a.browser.find_element(By.XPATH,
                                                          '//*[@id="app"]/div[1]/span[2]/div[1]/span/div[1]/div/div/div')
                    except:
                        checkvalid = 'none'
                    if checkvalid == 'none':
                        a.send_msg()
                        sleep(8)

                except:
                    sleep(10)

            #else:
                #pass
            progress_value = progress_value + start_value
            root.update()
        concluido_msg = tk.Label(root, text='Process finished!', fg='green')
        concluido_msg.place(x=280, y=860, width=200, height=25)
        root.update()

    def start_process(self):
        a.make_chrome_browser()
        a.start_browser()
        a.login_prepona()
        a.verificar_dispositivo()
        a.selecionar_filtros()
        a.verificar_candidatos()
        a.pular_pagina()
        a.pular_pagina()
        a.pular_pagina()
        a.open_cert()
        a.cert_login()
        a.get_celular()
        a.check_situation()
        a.browser.get('https://web.whatsapp.com/')
        a.check_if_open()
        a.send_wpp()
        a.criar_planilha()
        a.browser.quit()



a = App()

root = Tk()
root.title("FGV Cert")
root.iconbitmap('C:/CertPrepona/fgv-logo-0.ico')
root.geometry('800x900')
#Title
image_logo = Image.open("C:/CertPrepona/fgv-logo-1.png")
fgv_logo = ImageTk.PhotoImage(image_logo)
label1 = tk.Label(image=fgv_logo)
label1.image = fgv_logo
label1.place(x=50,y=0,width=700,height=200)
#Imagechnacib
image2 = Image.open("C:/CertPrepona/chnacib.jpg")
image_logo = ImageTk.PhotoImage(image2)
label2 = tk.Label(image=image_logo)
label2.image = image_logo
label2.place(x=0,y=700,width=200,height=150)
#Calendário
cal = Calendar(root, selectmode='day', year=year_today, month=month_today, day=day_today)
cal.place(x=270,y=220,width=250,height=200)
cal_get = Button(root,text='Get date',command=a.data_get)
cal_get.place(x=340,y=435,width=100,height=30)

#Começar processo
start_btn = Button(root,text='Start',command=a.start_process)
start_btn.place(x=280,y=830,width=200,height=30)

#Botão importsheet
planilha_btn = Button(root,text='Import sheet',command=a.importar_planilha)
planilha_btn.place(x=280,y=600,width=200,height=30)

#Assinatura
assinatura = tk.Label(root, text='github.com/chnacib', fg='black')
assinatura.place(x=0,y=850,width=200,height=30)

root.mainloop()
