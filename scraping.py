import re
import openpyxl
import yagmail
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from time import sleep
from tqdm import tqdm


def verifica_email(email):
    padrao = re.compile(r'^[a-zA-Z0-9]+@gmail\.com$')

    if re.fullmatch(pattern=padrao, string=email):
        return True
    else:
        return False
    
def enviar_email(destino):
    yag = yagmail.SMTP('Testesdepython@gmail.com', 'jdqx yguk zapp sgkd')

    # Destinatário e assunto do email
    destinatario = destino
    assunto = 'Busca feita'

    # Corpo do email
    corpo_email = "Segue o arquivo com os preços dos produtos atualizados."

    # Anexos
    anexos = ['web_scraping.xlsx']  # Substitua pelo caminho do seu arquivo

    # Enviar o email
    yag.send(
        to=destinatario,
        subject=assunto,
        contents=corpo_email,
        attachments=anexos
        )


email = str(input("Seu email para o envio do arquivo: "))

while verifica_email(email) != True:
    print("Email inváldo! Confira se está no formato gmail ou houve erro de dígito.\n")
    email = str(input("Seu email para o envio do arquivo: "))

print("\n\033[1;38mIniciando...\033[m\n")

planilha = openpyxl.Workbook()
planilha.create_sheet('Valores')
tabela = planilha['Valores']
tabela.append(['Produto', 'Valor'])

site = 'https://www.amazon.com.br/'
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
navegador.get(site)
sleep(2)

navegador.find_element(By.XPATH, '//*[@id="twotabsearchtextbox"]').send_keys('aliança casal')
navegador.find_element(By.XPATH, '//*[@id="nav-search-submit-button"]').click()
sleep(1)

while True:
    soup = BeautifulSoup(navegador.page_source, 'html.parser')
    itens = soup.find_all('div', attrs={'class': 'a-section a-spacing-small puis-padding-left-micro puis-padding-right-micro'})
    
    for item in tqdm(itens):
        produtos = item.find('h2', attrs={'class': 'a-size-mini a-spacing-none a-color-base s-line-clamp-2'})
        reais = item.find('span', attrs={'class': 'a-price-whole'})
        centavos = item.find('span', attrs={'class': 'a-price-fraction'})

        if produtos and reais and centavos:    
            tabela.append([produtos.text, f'R$ {reais.text}{centavos.text}'])

    proxima = soup.find('a', attrs={'class': 's-pagination-item s-pagination-next s-pagination-button s-pagination-separator'})
    
    if proxima:
        print("\033[32mPróxima Página\033[m")
        sleep(2)
        navegador.get(site + str(proxima['href']))
    else:
        navegador.quit()
        planilha.save('web_scraping.xlsx')
        enviar_email(email)
        break

print("\033[32m\nFim da busca! Planilha feita.\nEmail enviado com sucesso!\033[m")