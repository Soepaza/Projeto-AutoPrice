import pandas as pd
import requests

df = pd.read_excel("buscas.xlsx")
# print(df)

li = df["Nome"].to_list()
print(li)

for item in li:
    url = requests.get(
        f'https://www.buscape.com.br/search?q={item.replace(" ", "%")}')
    print(url.content)
    break

print(item[1])