import openpyxl
import requests
from bs4 import BeautifulSoup as bs


def leap(year):
    if (year % 4 == 0 and year % 100 != 0) or year % 400 == 0:
        return True
    else:
        return False


def create_html_list(year):
    # https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_28_fevralya_2022_goda.html
    html_list = list([])
    months = ['dekabrya', 'noyabrya', 'oktyabrya', 'sentyabrya', 'avgusta', 'iyulya', 'iyunya', 'maya', 'aprelya',
              'marta', 'fevralya', 'yanvarya']
    months = list(reversed(months))
    last_days = ['31', '28', '31', '30', '31', '30', '31', '31', '30', '31', '30', '31']
    if leap(year):
        last_days[1] = '29'
    for i in range(len(months)):
        first_part = 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_'
        second_part = '1'
        third_part = '_po_'
        fourth_part = last_days[i]
        fifth_part = '_'
        sixth_part = months[i]
        seventh_part = '_'
        eight_part = str(year)
        nine_part = '_goda.html'
        html_list.append(
            first_part + second_part + third_part + fourth_part + fifth_part + sixth_part + seventh_part + eight_part + nine_part)
    return html_list


def get_duty(url):
    headers = {'user-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:72.0) Gecko/20100101 Firefox/72.0',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
               }
    response = requests.get(url, headers=headers)
    soup = bs(response.text, 'html.parser')
    soup = soup.findAll('p')
    rows = list(map(lambda x: x.text, soup))
    duty = 0
    for i in range(1, len(rows)):
        if rows[i - 1] == 'нефть сырая':
            duty = float(rows[i].replace(',', '.'))
            break
    return duty


html_list = create_html_list(2022)
duties = list([])
for url in html_list:
    duties.append(get_duty(url))

wb = openpyxl.load_workbook('Приложение 1.xlsx')
sheet = wb['Company ABC_факт_НДПИ (Platts)']
for i in range(len(duties)):
    sheet[f"{chr(67 + i)}{7}"].value = duties[i]
wb.save("Приложение 1.xlsx")
wb.close()

