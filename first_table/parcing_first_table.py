from bs4 import BeautifulSoup as bs
import urllib.request as urllib2
from urllib.request import Request, urlopen
from urllib.request import urlopen
import numpy as np
import pandas as pd
import openpyxl
def spread_quotes(url):
    req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    web_byte = urlopen(req).read()

    webpage = web_byte.decode('utf-8')
    soup = bs(webpage, "html.parser")
    quotes = soup.find_all('td', class_=['datatable_cell__LJp3C datatable_cell--align-end__qgxDQ datatable_cell--up__hIuZF text-right text-sm font-normal leading-5 align-middle min-w-[77px] text-[#007C32]',
                                     'datatable_cell__LJp3C datatable_cell--align-end__qgxDQ datatable_cell--down___c4Fq text-right text-sm font-normal leading-5 align-middle min-w-[77px] text-[#D91400]'])
    quotes = np.array([float(x.text.replace(',', '.')) for x in quotes])
    return quotes

url = 'https://ru.investing.com/commodities/brent-wti-crude-spread-futures-historical-data'
quotes = spread_quotes(url)
print(quotes)

wb = openpyxl.load_workbook('Приложение 1.xlsx')
sheet = wb['Анализ_БК+ББ']

print(sheet)
for i in range(len(quotes)):
    # sheet[] = quotes[i]
    sheet[f"C{4+i}"].value = quotes[i]
wb.save("Приложение 1.xlsx")
wb.close()



