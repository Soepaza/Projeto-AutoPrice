import pandas as pd
import requests
from bs4 import BeautifulSoup
from pprint import pprint
import win32com.client as win32
import pathlib

produtos = pd.read_excel('buscas.xlsx')
lista_nomes = produtos['Nome'].to_list()

all_data = []
url_original = 'https://www.buscape.com.br'

for nome in lista_nomes:
    url = f'https://www.buscape.com.br/search?q={nome}'
    print(f"Buscando preços para o produto: {nome}, {url}")
    
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')

    link_url = soup.find_all('a', class_='ProductCard_ProductCard_Inner__gapsh')

    results_store = soup.find_all('h3')
    storeNames = [resultstore.string for resultstore in results_store]
    # print(storeNames)

    results_prices = soup.find_all('p', class_='Text_Text__ARJdp Text_MobileHeadingS__HEz7L')
    storePrice = [resultprice.string for resultprice in results_prices]
    # print(storePrice)

    lista_href = []
    for link in link_url:
        href = link.get('href')
        if 'https://www.buscape.com.br' not in  href:
            href = url_original + href
            lista_href.append(href)
    print(f'printando listas de referencias: {lista_href}')

    product_data = []
    for i in range(len(storeNames)):
        product_data.append({'Nome': nome, 'Loja': storeNames[i], 'Preço': storePrice[i], 'URL': lista_href[i]})

    
    product_data = sorted(product_data, key=lambda x: float(x['Preço'].replace('R$', '').replace(',', '').strip()))
    all_data.extend(product_data)

tabela_precos = pd.DataFrame(all_data)
print(tabela_precos)

path = r'C:\Users\Home\Desktop\Projetos Hashtag\Projeto2 - Automaçao Web\Projeto 2\produtosAchados.xlsx'
tabela_precos.to_excel(path, index=False)

caminho_arquivo = pathlib.Path('produtosAchados.xlsx')

outlook= win32.Dispatch("outlook.application")
email = outlook.CreateItem(0)

email.To = 'alison.spza+Teste@hotmail.com'
email.Subject = 'Preços encontrados'

attachment = pathlib.Path.cwd() / caminho_arquivo
email.Attachments.Add(str(attachment))

email.Send()



#organizar o projeto no github


#organizar por preços os produtos ////////////// / FEITO
#filtrar preco minimo e preco maximo    ///////// Nao feito
#mandar email automatico   /////////// Feito

#