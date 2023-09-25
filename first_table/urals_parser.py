from bs4 import BeautifulSoup as bs
import urllib.request as urllib2
from urllib.request import Request, urlopen
from urllib.request import urlopen
import numpy as np
import pandas as pd
import openpyxl

# "https://www.economy.gov.ru/material/departments/d12/konyunktura_mirovyh_tovarnyh_rynkov/o_sredney_cene_na_neft_sorta_yurals_za_yanvar_2022_goda.html"
url = ["https://www.economy.gov.ru/material/departments/d12/konyunktura_mirovyh_tovarnyh_rynkov/o_sredney_cene_na_neft_sorta_yurals_za_", "_2022_goda.html"]
months = ['fevral', 'mart','aprel','may','iyun','iyul', 'avgust', 'sentyabr', 'oktyabr', 'noyabr', 'dekabr']
prices = []
for month in months:
    print(month)
    url_link = url[0]+month+url[1]
    req = Request(url_link, headers={'User-Agent': 'Mozilla/5.0'})
    web_byte = urlopen(req).read()

    webpage = web_byte.decode('utf-8')
    soup = bs(webpage, "html.parser")
    prices.append(float([x for x in soup.findAll('p') if "США за баррель" in x.text][0].text.split()[0].replace(",", ".")))

wb = openpyxl.load_workbook('Приложение 1.xlsx')
sheet = wb['Компания 1_факт_НДПИ (Platts)']

print(sheet)
for i in range(len(prices)):
    # sheet[] = quotes[i]
    sheet[f"{chr(68+i)}{14}"].value = prices[i]
    sheet[f"{chr(68+i)}{171}"].value = prices[i]
wb.save("Приложение 1.xlsx")
wb.close()


