import pandas as pd
import requests
from bs4 import BeautifulSoup

produtos = pd.read_excel('buscas.xlsx')
lista_nomes = produtos['Nome'].to_list()

all_data = []  # Lista para armazenar os dados de todos os produtos

for nome in lista_nomes:
    url = f'https://www.buscape.com.br/search?q={nome}'
    
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')

    results_store = soup.find_all('h3')
    storeNames = [resultstore.string for resultstore in results_store]

    results_prices = soup.find_all('p', class_='Text_Text__ARJdp Text_MobileHeadingS__HEz7L')
    storePrice = [resultprice.string for resultprice in results_prices]

    # Adicionar dados a all_data para cada loja e preço
    for store, price in zip(storeNames, storePrice):
        all_data.append({'Produto': nome, 'Loja': store, 'Preço': price})

# Criar DataFrame final
tabela_precos = pd.DataFrame(all_data)
print(tabela_precos)