import re
import openpyxl
import yagmail
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from time import sleep


def verifica_email(email):
    padrao = re.compile(r'^[a-zA-Z0-9]+@gmail\.com$')

    if re.fullmatch(pattern=padrao, string=email):
        return True
    else:
        return False

email = str(input("Seu email para o envio do arquivo: "))

while verifica_email(email) != True:
    print("Email inváldo! Confira se está no formato gmail ou houve erro de dígito.\n")
    email = str(input("Seu email para o envio do arquivo: "))

print("Iniciando...\n")

planilha = openpyxl.Workbook()
planilha.create_sheet('Valores')
tabela = planilha['Valores']
tabela.append(['Produto', 'Valor'])

site = 'https://www.amazon.com.br/'
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
navegador.get(site)
sleep(2)

navegador.find_element(By.XPATH, '//*[@id="twotabsearchtextbox"]').send_keys('smartphones samsung')
navegador.find_element(By.XPATH, '//*[@id="nav-search-submit-button"]').click()
sleep(1)

while True:
    soup = BeautifulSoup(navegador.page_source, 'html.parser')
    itens = soup.find_all('div', attrs={'class': 'a-section a-spacing-small puis-padding-left-small puis-padding-right-small'})
    
    for item in itens:
        produtos = item.find('span', attrs={'class': 'a-size-base-plus a-color-base a-text-normal'})
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
        yag = yagmail.SMTP('Testesdepython@gmail.com', 'jdqx yguk zapp sgkd')

        # Destinatário e assunto do email
        destinatario = email
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

        break

print("Fim da busca! Planilha feita.")
print("Email enviado com sucesso!")