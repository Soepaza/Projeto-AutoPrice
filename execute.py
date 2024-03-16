import pandas as pd
import requests
from bs4 import BeautifulSoup
import win32com.client as win32
import pathlib
import os


def enviar_email(destinatario, assunto, corpo, anexo=None):
    outlook = win32.Dispatch("outlook.application")
    email = outlook.CreateItem(0)
    email.To = destinatario
    email.Subject = assunto
    email.Body = corpo

    if anexo:
        attachment = pathlib.Path.cwd() / anexo
        email.Attachments.Add(str(attachment))

    try:
        email.Send()
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f'Erro ao enviar e-mail: {e}')


def buscar_precos_e_enviar_email():
    produtos = pd.read_excel('buscas.xlsx')
    lista_nomes = produtos['Nome'].to_list()

    all_data = []
    url_original = 'https://www.buscape.com.br'

    for nome in lista_nomes:
        url = f'https://www.buscape.com.br/search?q={nome}'
        print(f"Buscando preços para o produto: {nome}, {url}")

        page = requests.get(url)
        soup = BeautifulSoup(page.text, 'html.parser')

        link_url = soup.find_all(
            'a', class_='ProductCard_ProductCard_Inner__gapsh')

        results_store = soup.find_all('h3')
        storeNames = [resultstore.string for resultstore in results_store]

        results_prices = soup.find_all(
            'p', class_='Text_Text__ARJdp Text_MobileHeadingS__HEz7L')
        storePrice = [resultprice.string for resultprice in results_prices]

        lista_href = []
        for link in link_url:
            href = link.get('href')
            if 'https://www.buscape.com.br' not in href:
                href = url_original + href
                lista_href.append(href)
        # print(f'printando listas de referencias: {lista_href}')

        product_data = []
        for i in range(len(storeNames)):
            product_data.append(
                {'Nome': nome, 'Loja': storeNames[i], 'Preço': storePrice[i], 'URL': lista_href[i]})

        product_data = sorted(product_data, key=lambda x: float(
            x['Preço'].replace('R$', '').replace(',', '').strip()))
        all_data.extend(product_data)

    tabela_precos = pd.DataFrame(all_data)

    path_saida = pathlib.Path('output') / 'produtosAchados.xlsx'
    try:
        tabela_precos.to_excel(path_saida, index=False)
    except FileNotFoundError:
        pathlib.Path('output').mkdir(parents=True, exist_ok=True)
        tabela_precos.to_excel(path_saida, index=False)

    caminho_arquivo = pathlib.Path.cwd() / path_saida

    enviar_email('soetestes+1@outlook.com', 'Preços encontrados',
                 'Confira os preços encontrados nos anexos.', caminho_arquivo)

    enviar_email('soetestes+2@outlook.com', 'New! Novos preços encontrados',
                 'Novo!!! Confira os novos preços', caminho_arquivo)


buscar_precos_e_enviar_email()
