#Chromedriver executable needs to be in PATH. See https://chromedriver.chromium.org/downloads

import os

#os.system('pip install selenium')
#os.system('pip install unidecode')
#os.system('pip install arrow')
#os.system('pip install pandas')

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import *
from unidecode import unidecode
import re
import arrow
import time
from time import sleep
import datetime
from datetime import date
import pandas as pd
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def line(size=45):
    return '-' * size

def header(txt):
    print(line())
    print(txt.center(54))
    print(line())

def valid_date(datestring):
    data_atual = arrow.now().format('YYYY-MM-DD')
    formatted_data_atual = time.strptime(data_atual, "%Y-%m-%d")
    formatted_date2 = time.strptime(datestring, "%Y-%m-%d")
    try:
        datetime.datetime.strptime(datestring, '%Y-%m-%d')
        if formatted_data_atual <= formatted_date2:
            return True
    except ValueError:
        return False

def highlight_adversas(s):
    color_red = 'red'
    color_yellow = 'yellow'
    color_green = 'green'
    if s >= 15:
        return 'background-color: %s' % color_red
    elif s < 15 and s > 0:
        return 'background-color: %s' % color_yellow
    elif s <= 0:
        return 'background-color: %s' % color_green


header(f'\033[33mHotel Prices Webscraper Decolar x Booking\033[m')

# IDs destinos
  #Aracaju: 180
  #Arraial dAjuda: 339
  #Balneario Camboriu: 192594
  #Belo Horizonte: 701
  #Brasilia: 926
  #Buzios: 1077
  #Cabo Frio: 1227
  #Caldas Novas: 1364
  #Campinas: 1467
  #Campos do Jordao: 6116
  #Canela: 6064
  #Curitiba: 1595
  #Florianopolis: 2261
  #Fortaleza: 2302
  #Foz do Iguacu: 3072
  #Gramado: 2611
  #Ilhabela: 3027
  #Ilheus: 3158
  #Joao Pessoa: 3399
  #Maceio: 4430
  #Morro de Sao Paulo: 4678
  #Natal: 4971
  #Porto Alegre: 5822
  #Porto de Galinhas: 5648
  #Porto Seguro: 881
  #Praia de Pipa: 5875
  #Recife: 6322
  #Rio de Janeiro: 6381
  #Salvador: 7018
  #Sao Paulo: 6574

# Input parâmetros
print('Definição de parâmetros\n')

print('Opções de destinos:')                
lista_destinos = ['Aracaju', 'Arraial dAjuda', 'Balneario Camboriu', 'Belo Horizonte', 'Brasilia', 'Buzios', 'Cabo Frio', 'Caldas Novas',
                  'Campinas', 'Campos do Jordao', 'Canela', 'Curitiba', 'Florianopolis', 'Fortaleza', 'Foz do Iguacu', 'Gramado',
                  'Ilhabela', 'Ilheus', 'Joao Pessoa', 'Maceio', 'Morro de Sao Paulo', 'Natal', 'Porto Alegre', 'Porto de Galinhas',
                  'Porto Seguro', 'Praia de Pipa', 'Recife', 'Rio de Janeiro', 'Salvador', 'Sao Paulo']

lista_ids = ['180', '339', '192594', '701', '926', '1077', '1227', '1364', '1467', '6116', '6064', '1595', '2261', '2302', '3072', '2611', '3027', '3158', '3399', '4430', '4678', '4971', '5822', '5648', '881', '5875', '6322', '6381', '7018', '6574']

opt = 1
for d in lista_destinos:
    print(f'{opt} - \033[36m{d}\033[m')
    opt += 1
print()
destInput = input('Digite sua opção de destino: ')
while True:  
    try:
        destInt = int(destInput)
        if destInt > 0 and destInt <= 30:
            nome_destino = lista_destinos[(destInt -1)]
            id_destino_decolar = lista_ids[(destInt -1)]
            break
        else:
            destInput = input('Escolha uma opção válida e digite novamente: ')
    except:
        destInput = input('Escolha uma opção válida e digite novamente: ')
        continue

cinInput = input('Digite a data de check-in [AAAA-MM-DD]: ')
while True:
    try:
        cinCheck = valid_date(cinInput)
        if cinCheck:
            checkin = cinInput
            break
        else:
            cinInput = input('Digite uma data de check-in válida: ')
    except:
        cinInput = input('Digite uma data de check-in válida: ')
        continue        

coutInput = input('Digite a data de check-out [AAAA-MM-DD]: ')
while True:
    try:
        coutCheck = valid_date(coutInput)
        if coutCheck and coutInput > checkin:
            checkout = coutInput
            break
        else:
            coutInput = input('Digite uma data de check-out válida: ')
    except:
        cinInput = input('Digite uma data de check-out válida: ')
        continue

paxInput = input('Digite o número de hóspedes [adultos]: ')
while True:
    try:
        paxInt = int(paxInput)
        if paxInt >= 1 and paxInt <= 4:
            pax = paxInt
            break
        else:
            paxInput = input('Digite um número de hóspedes entre 1 e 4: ')
    except:
        paxInput = input('Digite um número de hóspedes entre 1 e 4: ')
        continue

emailReport = input('Digite seu endereço de e-mail para receber o relatório: ')
patternMail = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
while True:
    if re.match(patternMail, emailReport):
        break
    else:
        emailReport = input("Digite um endereço de e-mail válido: ")

# Definições Chromedriver
chrome_options = Options()
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument('--lang=pt-BR')
chrome_options.add_argument('--disable-notifications')

# Scrape Booking
print()
print(f'\033[36mIniciando coleta de dados Booking...\033[m')

nome_hoteis_bkg = []
preco_hoteis_bkg = []

class ScrappyBkg:

    def iniciar(self):
        self.raspagem_de_dados()

    def raspagem_de_dados(self):
        cin = {"year": checkin[0:4],
               "month": checkin[5:7],
               "day": checkin[8:10]}
        cout = {"year": checkout[0:4],
                "month": checkout[5:7],
                "day": checkout[8:10]}
        chd = "0"
        rooms = "1"

        self.driver = webdriver.Chrome(options=chrome_options)
        self.link = f'https://www.booking.com/searchresults.pt-br.html?label=gen173nr-1FCAEoggI46AdIM1gEaCCIAQGYAS24ARnIAQzYAQHoAQH4AQuIAgGoAgO4At_SvIsGwAIB0gIkNzI1ZmZlYWEtZWRiYS00OWZmLWI5MzItZjYxNWMwYmY3N2U02AIG4AIB&lang=pt-br&sid=7b8aa653793caf99eb3d54cf87f15296&sb=1&sb_lp=1&src=index&src_elem=sb&error_url=https%3A%2F%2Fwww.booking.com%2Findex.pt-br.html%3Flabel%3Dgen173nr-1FCAEoggI46AdIM1gEaCCIAQGYAS24ARnIAQzYAQHoAQH4AQuIAgGoAgO4At_SvIsGwAIB0gIkNzI1ZmZlYWEtZWRiYS00OWZmLWI5MzItZjYxNWMwYmY3N2U02AIG4AIB%3Bsid%3D7b8aa653793caf99eb3d54cf87f15296%3Bsb_price_type%3Dtotal%3Bsig%3Dv1dXShAKR_%26%3B&ss={nome_destino}&is_ski_area=0&ssne={nome_destino}&ssne_untouched={nome_destino}&dest_id=-&dest_type=city&checkin_year={cin["year"]}&checkin_month={cin["month"]}&checkin_monthday={cin["day"]}&checkout_year={cout["year"]}&checkout_month={cout["month"]}&checkout_monthday={cout["day"]}&group_adults={pax}&group_children={chd}0&no_rooms={rooms}&b_h4u_keep_filters=&from_sf=1&order=review_score_and_price'
        self.lista_nome_hoteis = []
        self.lista_preco_hoteis = []
        self.driver.get(self.link)
        #self.driver.find_element(By.XPATH, '//*[@id="ajaxsrwrap"]/div[2]/div/div/div[2]/ul/li[3]/a').click()
        sleep(3)
        
        for p in range(20): # Passar como parâmetro nº de páginas a percorrer. Nº > total, percorre todas.
            item = 1
            lista_length = self.driver.find_elements(By.CLASS_NAME, 'e75f1d9859')
            lista_max = len(lista_length)
            for i in range(lista_max):
                c = 1
                desc = 1
                while c < lista_max:
                    try:
                        lista_nomes = self.driver.find_elements(By.XPATH,
                            f'//*[@id="search_results_table"]/div[1]/div/div/div/div[5]/div[{item}]/div[1]/div[2]/div/div[1]/div/div[1]/div/div[1]/div/h3/a/div[1]')
                        self.lista_nome_hoteis.append(lista_nomes[0].text)
                        sleep(1)
                        lista_precos = self.driver.find_elements(By.XPATH,
                            f'//*[@id="search_results_table"]/div[1]/div/div/div/div[5]/div[{item}]/div[1]/div[2]/div/div[3]/div/div[2]/div/div[1]/div/div/div/div/div[2]/span')                        
                        while len(lista_precos) == 0:
                            try:
                                lista_precos = self.driver.find_elements(By.XPATH,
                                    f'//*[@id="search_results_table"]/div[1]/div/div/div/div[5]/div[{desc}]/div[1]/div[2]/div/div[4]/div/div[2]/div/div[1]/div/div/div/div/div[2]/span[2]')
                                desc += 1
                            except:
                                desc += 1
                        self.lista_preco_hoteis.append(lista_precos[0].text)          
                        sleep(1)
                        item += 1
                    except:
                        item += 1
                        c += 1
            try:
                botao_proximo = self.driver.find_element(By.XPATH,
                    '//*[@id="search_results_table"]/div[1]/div/div/div/div[6]/div[2]/nav/div/div[3]/button')
                botao_proximo.click()
                print(f'\033[36mNavegando...\033[m')
                sleep(5)
            except:
                self.driver.quit()
        
        print(f'\033[36mRaw data:\033[m')
        print(self.lista_nome_hoteis)
        print(self.lista_preco_hoteis)

        for nome in self.lista_nome_hoteis:
            nome_hoteis_bkg.append(nome)
        for preco in self.lista_preco_hoteis:
            preco_hoteis_bkg.append(preco)           

start = ScrappyBkg()
start.iniciar()

print(f'\033[36mTratando dados...\033[m')
# Tratamento formato
nome_hoteis_bkg_temp = list(pd.Series(nome_hoteis_bkg).str.upper())
nome_hoteis_bkg_fmtg = []
for nome in nome_hoteis_bkg_temp:
    nome_hoteis_bkg_fmtg.append(unidecode(nome))

preco_hoteis_bkg_nofloat = []
for x in preco_hoteis_bkg:
    item = x
    for y in ["."]:
        item = item.replace(y, "")
    preco_hoteis_bkg_nofloat.append(item)

preco_hoteis_bkg_nosign = []
for x in preco_hoteis_bkg_nofloat:
    item = x
    for y in ["R$ "]:
        item = item.replace(y, "")
    preco_hoteis_bkg_nosign.append(item)

preco_hoteis_bkg_nosign = pd.to_numeric(preco_hoteis_bkg_nosign, errors='coerce', downcast='unsigned')

# Mesclagem nomes e preços em dicionário
data_hoteis_bkg = []
for el in zip(nome_hoteis_bkg_fmtg, preco_hoteis_bkg_nosign):
    data_hoteis_bkg.append(el)
data_hoteis_bkg_dict = dict(data_hoteis_bkg)

# Conversão dicionário para dataframe
df_data_hoteis_bkg = pd.DataFrame(list(data_hoteis_bkg_dict.items()),
                   columns=['Nome', 'Preço Booking'])

# Scrape Decolar
print()
print(f'\033[36mIniciando coleta de dados Decolar...\033[m')

nome_hoteis_desp = []
preco_hoteis_desp = []

class ScrappyDesp:

    def iniciar(self):
        self.raspagem_de_dados()

    def raspagem_de_dados(self):
        page = 1
        self.driver = webdriver.Chrome(options=chrome_options)
        self.link = f'https://www.decolar.com/accommodations/results/CIT_{id_destino_decolar}/{checkin}/{checkout}/{pax}?from=SB2&facet=city&searchId=2c403941-a9b6-457f-a17e-5f3131886e89&page={page}'
        self.lista_nome_hoteis = []
        self.lista_preco_hoteis = []
        self.driver.get(self.link)
        self.driver.find_element(By.XPATH, '//*[@id="lgpd-banner"]/div/a[2]').click()
        sleep(3)

        for p in range(20): # Passar como parâmetro nº de páginas a percorrer. Nº > total, percorre todas.
            item = 1
            lista_length = self.driver.find_elements(By.CLASS_NAME, 'accommodation-name')
            lista_max = len(lista_length)
            for i in range(lista_max):
                c = 1
                while c < lista_max:
                    try:
                        lista_nomes = self.driver.find_elements(By.XPATH,
                            f'/html/body/aloha-app-root/aloha-results/div/div/div/div/div[2]/aloha-list-view-container/div[2]/div[{item}]/aloha-cluster-container/div/div/div[1]/div/div[2]/div/aloha-cluster-accommodation-info-container/div[1]/span')
                        self.lista_nome_hoteis.append(lista_nomes[0].text)
                        sleep(1)
                        lista_precos = self.driver.find_elements(By.XPATH,
                            f'/html/body/aloha-app-root/aloha-results/div/div/div/div/div[2]/aloha-list-view-container/div[2]/div[{item}]/aloha-cluster-container/div/div/div[2]/aloha-cluster-pricebox-container/div/div[2]/div[1]/aloha-price-container/aloha-summary-price/div/span[2]')
                        self.lista_preco_hoteis.append(lista_precos[0].text)          
                        sleep(1)
                        item += 1
                    except:
                        item += 1
                        c += 1                        
            try:
                page += 1
                self.driver.get(f'https://www.decolar.com/accommodations/results/CIT_{id_destino_decolar}/{checkin}/{checkout}/{pax}?from=SB2&facet=city&searchId=2c403941-a9b6-457f-a17e-5f3131886e89&page={page}')
                print(f'\033[36mNavegando...\033[m')
                sleep(5)
            except:
                self.driver.quit()
        
        print(f'\033[36mRaw data:\033[m')
        print(self.lista_nome_hoteis)
        print(self.lista_preco_hoteis)

        for nome in self.lista_nome_hoteis:
            nome_hoteis_desp.append(nome)
        for preco in self.lista_preco_hoteis:
            preco_hoteis_desp.append(preco)

start = ScrappyDesp()
start.iniciar()

print(f'\033[36mTratando dados...\033[m')
# Tratamento formato
nome_hoteis_desp_temp = list(pd.Series(nome_hoteis_desp).str.upper())
nome_hoteis_desp_fmtg = []
for nome in nome_hoteis_desp_temp:
    nome_hoteis_desp_fmtg.append(unidecode(nome))

preco_hoteis_desp_nofloat = []
for x in preco_hoteis_desp:
    item = x
    for y in ["."]:
        item = item.replace(y, "")
    preco_hoteis_desp_nofloat.append(item)

preco_hoteis_desp_num = pd.to_numeric(preco_hoteis_desp_nofloat, errors='coerce')

# Mesclagem nomes e preços em dicionário
data_hoteis_desp = []
for el in zip(nome_hoteis_desp_fmtg, preco_hoteis_desp_num):
    data_hoteis_desp.append(el)
data_hoteis_desp_dict = dict(data_hoteis_desp)

# Conversão dicionário para dataframe
df_data_hoteis_desp = pd.DataFrame(list(data_hoteis_desp_dict.items()),
                   columns=['Nome', 'Preço Decolar'])

print(f'\033[36mFinalizando tratamento dos dados e gerando comparativo...\033[m')
# Merge data
data_merge = pd.merge(df_data_hoteis_bkg, df_data_hoteis_desp, on=['Nome'], how='left')

# Limpeza NaN
data_merge_notna = data_merge.dropna()

# Comparação de preços
df_rate_comparison = data_merge_notna

df_rate_comparison['% Diferença'] = 0
for el in df_rate_comparison:
    dif = (df_rate_comparison['Preço Decolar'] / df_rate_comparison['Preço Booking'] -1 ) * 100
    df_rate_comparison['% Diferença'] = round(dif, 0)

# Highlight adversas
reporte_adversas = df_rate_comparison.style.applymap(highlight_adversas, subset=pd.IndexSlice[:, ['% Diferença']])

print(f'\033[36mBaixando relatório...\033[m')
# Conversão dataframe para XLSX
writer = pd.ExcelWriter('Reporte Paridade Booking - {}.xlsx'.format(nome_destino))
reporte_adversas.to_excel(writer, f'IN_{checkin} - OUT_{checkout}', index=False)
writer.save()

# % adversas
medicoes = 0
for m in df_rate_comparison['% Diferença']:
    medicoes += 1

adversas = 0
for a in df_rate_comparison['% Diferença']:
    if a > 0:
        adversas += 1

percent_adversas = (adversas / medicoes) * 100

# Calcula o nº de noites
from datetime import datetime
d2 = datetime.strptime(checkout, '%Y-%m-%d')
d1 = datetime.strptime(checkin, '%Y-%m-%d')
n_diarias = abs((d2 - d1).days)

print(f'\033[36mEnviando relatório por e-mail...\033[m')
# Dispara e-mail
anexo = "Reporte Paridade Booking - {}.xlsx".format(nome_destino)

def sendMail():
    try:
        subject = "Reporte Paridade Booking - {}".format(nome_destino)
        body = f"Destino: {nome_destino}\n\nCheck-in: {checkin}\nCheck-out: {checkout}\nLOS: {n_diarias}\nPax: {pax}\n\nPercentual adversas: {percent_adversas:.2f}%"
        sender_email = "daniel.farias@decolar.com"
        password = ""
        receiver_email = emailReport
        cc_email = "dan.tfarias@gmail.com"

        message = MIMEMultipart()

        message["From"] = sender_email
        message["To"] = receiver_email
        message["Cc"] = cc_email
        message["Subject"] = subject

        message.attach(MIMEText(body, "plain"))

        filename = anexo

        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)

        part.add_header(
            "Content-Disposition",
            f"attachment; filename = {filename}",
        )

        message.attach(part)
        text = message.as_string()

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, text)
        
        print(f'\033[36mE-mail enviado com sucesso!\033[m')
    except:
        print(f'\033[31mErro ao enviar e-mail\033[m')


sendMail()