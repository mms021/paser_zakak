from selenium import webdriver
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

from bs4 import BeautifulSoup
from requests.api import request
from fake_headers import Headers
import requests as req

from openpyxl.styles import Font 
from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook

import tkinter as tk                
from tkinter import font  as tkfont 
from tkinter import messagebox ,ttk
import os , sys , asyncio ,  json 

header = Headers( browser="chrome", os="win",  headers=True  )

def elitan_parser(nom, qtw=100):
    password = {"username":"info@delta-pribor.ru","password":"hjvfirf61" ,"user-agent":"Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:94.0) Gecko/20100101 Firefox/94.0"}
    url = f"https://www.elitan.ru/price/index.php?find={nom.upper()}&delay=-1&mfg=all&seenform=y"
    option = webdriver.FirefoxOptions()
    # убирает флажок что автоматизированное ПО управляет браузером
    option.set_preference("dom.webdriver.enabled", False)
    # подмена user-agent
    option.set_preference("general.useragent.override", password["user-agent"])

    driver=webdriver.Firefox(options=option, executable_path=GeckoDriverManager().install())
    driver.get(url )
    el = WebDriverWait(driver ,timeout = 20).until(lambda d: d.find_elements_by_xpath('//*[@id="search_index"]/table/tbody/tr[6]/td'))
    
    print('d')
    #driver.close ()

def electronshik_parser(nom, qtw=100):
    url =  f"https://www.electronshik.ru/item/VISHAY/{nom.upper()}"
    driver=webdriver.Firefox(executable_path=GeckoDriverManager().install())
    driver.get(url)
    el = WebDriverWait(driver ,timeout = 20).until(lambda d: d.find_elements_by_xpath('//*[@id="dms-json"]'))
    data = json.loads(el[0].get_attribute("textContent"))
    clean_data = []
    #offer-header-dt tc ibb-m-d
    name = driver.find_element_by_class_name('item-page_name')
    for  i in range(1, len(data)):
        step = data[f'{i}'] 
        lev_data = []
        lev_data.extend(['Electronshik', url])
        lev_data.extend([name.text, step['max'] ,'' ]  )
        
        prise_data = []
        for prise in step["prices"]:
            pr = str(prise["min_qty"]) + " - "+ str(prise["max_qty"])
            prise_data.append([ pr , prise["price"] ])

        lev_data.append(prise_data)
        clean_data.append(lev_data)

    driver.close ()   
    return clean_data

def getchips_parser(nom , qtw=100):
    url = f"https://getchips.ru/rezultaty-poiska?input_field={nom.upper()}&no_cache=1&id=20&count_field={qtw}"
    s = req.Session()
    resp = s.get(url=url, headers=header.generate())
    soup = BeautifulSoup(resp.text, 'lxml')
    mydivs = soup.find_all("div", {"class": "result_price_data"})
    
    clean_data = []

    for parise_data in mydivs:
        data = json.loads(parise_data['rel'])
        lev_data = []
        lev_data.extend(['GetChips', url])
        lev_data.extend([data["title"] ,data["quantity"] , data['orderdays'] ])
        prise_data = []
        for prise in data["priceBreak"]:
            prise_data.append([prise["quantity"],prise["price"] ])

        lev_data.append(prise_data)
        clean_data.append(lev_data)
    
    return clean_data
    
def chipdip_parser(nom , qtw=100):
    #url = f'https://www.chipdip.ru/product/{nom.lower()}?from=suggest_product'
    url = f'https://www.chipdip.ru/search?searchtext={nom.upper()}'
    s = req.Session()
    resp = s.get(url=url, headers=header.generate())
    soup = BeautifulSoup(resp.text, 'lxml')
    prise_second  = soup.find_all("tr", {"class": "with-hover"}) 
    clean_data = []

    for i in prise_second:
        lev_data = []
        url_pr = f"https://www.chipdip.ru{i.find('a')['href']}"
        lev_data.extend(['ChipDip', url_pr ])
        lev_d = i.find_all("span", {"class": "nw"})
        lev_data.extend([ i.find('a').text ,  lev_d[1].text , lev_d[0].text ])
        prise_data = []

        elem_finde = i.find('input', {'class':'input input_qty'})['data-discounts']
        for a in elem_finde.split('],['):
            a1= a.replace("]", "").replace("[", "")
            prise_data.append(a1.split(','))

        lev_data.append(prise_data)
        clean_data.append(lev_data)

    return clean_data


def exel_file(name , data):
    wb = Workbook()
    wb.create_sheet(title = name, index = 0)
    sheet = wb[name]
    font = Font(size=12, bold=True)
    font2 = Font(size=8, italic=True)
    font3 = Font(bold=True)

    sheet.append([ 'Имя Сайта', "Сайт" , 'КОД', 'Колличество','' ])
    for i in data:
        sheet.append(i[:4])
        print(i[4])
        for h in i[4]:
            sheet.append(h)

    sheet.column_dimensions['A'].width = 10
    sheet.column_dimensions['B'].width = 10
    sheet.column_dimensions['C'].width = 40
    sheet.column_dimensions['D'].width = 6
    sheet.column_dimensions['E'].width = 6

    sheet['C2'].font = font
    sheet['C5'].font = font2
    sheet['C7'].font = font3

    wb.save('rezalt.xlsx')

name = 'RCS0805100RFKEA'

def main_func(name , nom = 100):
    data = chipdip_parser(name , nom) + (getchips_parser(name , nom)) + electronshik_parser(name, nom )
    exel_file(name , data)







#main_func(name)


print(electronshik_parser(name))
#electronshik_parser(n)
#getchips_parser(n)
#print(chipdip_parser(n))




