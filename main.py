from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui
import xlsxwriter
import os


open_browser = webdriver.Chrome()

open_browser.get('https://www.infomoney.com.br/ferramentas/cambio/')

pyautogui.sleep(0.5)

get_argentinian_peso_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[1]/td[3]')[0].text

get_argentinian_peso_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[1]/td[4]')[0].text

get_autralian_dollar_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[2]/td[3]')[0].text

get_autralian_dollar_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[2]/td[4]')[0].text

get_canadian_dollar_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[3]/td[3]')[0].text

get_canadian_dollar_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[3]/td[4]')[0].text

get_swiss_franc_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[4]/td[3]')[0].text

get_swiss_franc_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[4]/td[4]')[0].text

get_commecial_dollar_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[5]/td[3]')[0].text

get_commecial_dollar_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[5]/td[4]')[0].text

get_turism_dollar_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[6]/td[3]')[0].text

get_turism_dollar_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[6]/td[4]')[0].text

get_euro_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[7]/td[3]')[0].text

get_euro_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[7]/td[4]')[0].text

get_pound_sterling_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[8]/td[3]')[0].text

get_pound_sterling_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[8]/td[4]')[0].text

get_yen_purchase_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[9]/td[3]')[0].text

get_yen_selling_price = open_browser.find_elements(By.XPATH, '//*[@id="container_table"]/table/tbody/tr[9]/td[4]')[0].text

pyautogui.sleep(0.5)

open_browser.get('https://www.google.com/')

pyautogui.sleep(0.5)


search_mozambican_metical = open_browser.find_element(By.NAME, 'q').send_keys('Metical moçambicano')

pyautogui.sleep(0.5)

search_mozambican_metical = open_browser.find_element(By.NAME, 'q').send_keys(Keys.RETURN)

pyautogui.sleep(0.5)

get_mozambican_metical = open_browser.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div/span[1]').text

pyautogui.sleep(0.5)

clear_search = open_browser.find_element(By.NAME, 'q').send_keys(' ')

pyautogui.sleep(0.5)

pyautogui.press('tab')

pyautogui.press('enter')

search_chinese_yuan = open_browser.find_element(By.NAME, 'q').send_keys('Yuan chinês')

pyautogui.sleep(0.5)

search_chinese_yuan = open_browser.find_element(By.NAME, 'q').send_keys(Keys.RETURN)

get_chinese_yuan = open_browser.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div/span[1]').text

def convertStrinToFloat(currency):

    currency_replaced = currency.replace(',','.')

    currency_float = float(currency_replaced)

    return currency_float


file_path = 'C:\\Users\\lugan.costa\Desktop\\automation\\spreadsheets\\quoatition world currencies.xlsx'

currencies_spreadsheet = xlsxwriter.Workbook(file_path)

spreadsheet_1 = currencies_spreadsheet.add_worksheet()

spreadsheet_1.write('A1', 'Currency')
spreadsheet_1.write('A2', 'Peso argentino')
spreadsheet_1.write('A3', 'Dólar australiano')
spreadsheet_1.write('A4', 'Dólar canadense')
spreadsheet_1.write('A5', 'Franco suiço')
spreadsheet_1.write('A6', 'Dólar comercial')
spreadsheet_1.write('A7', 'Dólar turismo')
spreadsheet_1.write('A8', 'Euro')
spreadsheet_1.write('A9', 'Libra esterlina')
spreadsheet_1.write('A10', 'Iene japonês')
spreadsheet_1.write('A11', 'Metical moçambicano')
spreadsheet_1.write('A12', 'Yuan chinês')
spreadsheet_1.write('B1', 'Purchase Price')
spreadsheet_1.write('B2', convertStrinToFloat(get_argentinian_peso_purchase_price))
spreadsheet_1.write('B3', convertStrinToFloat(get_autralian_dollar_purchase_price))
spreadsheet_1.write('B4', convertStrinToFloat(get_canadian_dollar_purchase_price))
spreadsheet_1.write('B5', convertStrinToFloat(get_swiss_franc_purchase_price))
spreadsheet_1.write('B6', convertStrinToFloat(get_commecial_dollar_purchase_price))
spreadsheet_1.write('B7', convertStrinToFloat(get_turism_dollar_purchase_price))
spreadsheet_1.write('B8', convertStrinToFloat(get_euro_purchase_price))
spreadsheet_1.write('B9', convertStrinToFloat(get_pound_sterling_purchase_price))
spreadsheet_1.write('B10', convertStrinToFloat(get_yen_purchase_price))
spreadsheet_1.write('B11', convertStrinToFloat(get_mozambican_metical))
spreadsheet_1.write('B12', convertStrinToFloat(get_chinese_yuan))
spreadsheet_1.write('C1', 'Selling Price')
spreadsheet_1.write('C2', convertStrinToFloat(get_argentinian_peso_selling_price))
spreadsheet_1.write('C3', convertStrinToFloat(get_autralian_dollar_selling_price))
spreadsheet_1.write('C4', convertStrinToFloat(get_canadian_dollar_selling_price))
spreadsheet_1.write('C5', convertStrinToFloat(get_swiss_franc_selling_price))
spreadsheet_1.write('C6', convertStrinToFloat(get_commecial_dollar_selling_price))
spreadsheet_1.write('C7', convertStrinToFloat(get_turism_dollar_selling_price))
spreadsheet_1.write('C8', convertStrinToFloat(get_euro_selling_price))
spreadsheet_1.write('C9', convertStrinToFloat(get_pound_sterling_selling_price))
spreadsheet_1.write('C10', convertStrinToFloat(get_yen_selling_price))

currencies_spreadsheet.close()

os.startfile(file_path)


