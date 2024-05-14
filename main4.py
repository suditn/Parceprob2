import requests
import pandas as pd
import io
from bs4 import BeautifulSoup
from thefuzz import fuzz as f


cont = requests.get('https://www.vishay.com/en/product/34539/tab/designtools-ppg/').content
soup = BeautifulSoup(cont, "lxml")
list = []
for a in soup.findAll('a',href=True):

    list.append(a['href'])

for b in list:
    if b[-12:] == '_3dmodel.zip':
        print(b)

