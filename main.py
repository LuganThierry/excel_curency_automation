from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui
import xlsxwriter
import os
from currency_controller import CurrencyController 


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

pyautogui.sleep(0.5)

clear_search = open_browser.find_element(By.NAME, 'q').send_keys(' ')

pyautogui.sleep(0.5)

pyautogui.press('tab')

pyautogui.press('enter')

search_ruandan_franc = open_browser.find_element(By.NAME, 'q').send_keys('Franco ruandês')

pyautogui.sleep(0.5)

search_ruandan_franc = open_browser.find_element(By.NAME, 'q').send_keys(Keys.RETURN)

get_ruandan_franc = open_browser.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div/span[1]').text

pyautogui.sleep(0.5)

open_browser.get('https://www.infomoney.com.br/cotacoes/cripto/')

pyautogui.sleep(0.5)

get_cripto_bitcoin = open_browser.find_elements(By.XPATH, '//*[@id="ticker-datagrid-table-content"]/tr[2]/td[2]/span')[0].text

get_cripto_ethereum = open_browser.find_elements(By.XPATH, '//*[@id="ticker-datagrid-table-content"]/tr[3]/td[2]/span')[0].text

get_cripto_solana = open_browser.find_elements(By.XPATH, '//*[@id="ticker-datagrid-table-content"]/tr[6]/td[2]/span')[0].text

file_path = 'C:\\Users\\lugan.costa\Desktop\\automation\\spreadsheets\\quoatition world currencies.xlsx'

currencies_spreadsheet = xlsxwriter.Workbook(file_path)

cc = CurrencyController()

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
spreadsheet_1.write('A13', 'Franco ruandês')
spreadsheet_1.write('A14', 'Bitcoin')
spreadsheet_1.write('A15', 'Ethereum')
spreadsheet_1.write('A16', 'Solana')
spreadsheet_1.write('B1', 'Purchase Price')
spreadsheet_1.write('B2', cc.convertStringToFloat(get_argentinian_peso_purchase_price))
spreadsheet_1.write('B3', cc.convertStringToFloat(get_autralian_dollar_purchase_price))
spreadsheet_1.write('B4', cc.convertStringToFloat(get_canadian_dollar_purchase_price))
spreadsheet_1.write('B5', cc.convertStringToFloat(get_swiss_franc_purchase_price))
spreadsheet_1.write('B6', cc.convertStringToFloat(get_commecial_dollar_purchase_price))
spreadsheet_1.write('B7', cc.convertStringToFloat(get_turism_dollar_purchase_price))
spreadsheet_1.write('B8', cc.convertStringToFloat(get_euro_purchase_price))
spreadsheet_1.write('B9', cc.convertStringToFloat(get_pound_sterling_purchase_price))
spreadsheet_1.write('B10', cc.convertStringToFloat(get_yen_purchase_price))
spreadsheet_1.write('B11', cc.convertStringToFloat(get_mozambican_metical))
spreadsheet_1.write('B12', cc.convertStringToFloat(get_chinese_yuan))
spreadsheet_1.write('B13', cc.convertStringToFloat(get_ruandan_franc))
spreadsheet_1.write('B14', cc.convertCriptoToFloat(get_cripto_bitcoin))
spreadsheet_1.write('B15', cc.convertCriptoToFloat(get_cripto_ethereum))
spreadsheet_1.write('B16', cc.convertCriptoToFloat(get_cripto_solana))
spreadsheet_1.write('C1', 'Selling Price')
spreadsheet_1.write('C2', cc.convertStringToFloat(get_argentinian_peso_selling_price))
spreadsheet_1.write('C3', cc.convertStringToFloat(get_autralian_dollar_selling_price))
spreadsheet_1.write('C4', cc.convertStringToFloat(get_canadian_dollar_selling_price))
spreadsheet_1.write('C5', cc.convertStringToFloat(get_swiss_franc_selling_price))
spreadsheet_1.write('C6', cc.convertStringToFloat(get_commecial_dollar_selling_price))
spreadsheet_1.write('C7', cc.convertStringToFloat(get_turism_dollar_selling_price))
spreadsheet_1.write('C8', cc.convertStringToFloat(get_euro_selling_price))
spreadsheet_1.write('C9', cc.convertStringToFloat(get_pound_sterling_selling_price))
spreadsheet_1.write('C10', cc.convertStringToFloat(get_yen_selling_price))

currencies_spreadsheet.close()

os.startfile(file_path)

# TO DO
# Editar o tamanho das colunas do excel para visualizar as informações completas
# Adicionar data e hora da extração da informação
# Adicionar outras cotações
# Aprofundar a arquitetura: adicionar banco de dados onde se persiste os dados várias vezes
# Criação de API que gera o relatório a partir da cotação dos últimos 30 dias

