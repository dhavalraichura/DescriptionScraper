from selenium import webdriver
from selenium import *
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os,time
from selenium.webdriver import Keys
import pandas
import xlwt
from xlwt import Workbook
import pyautogui
PATH = ".\chromedriver.exe"
# chromeOptions = Options()
# chromeOptions.add_argument("--headless")
driver = webdriver.Chrome()
df = pandas.read_excel('.\Listings.xlsx')
wb = Workbook()
sheet1 = wb.add_sheet('Sheet1')
rownum = 0
values = df['CWI Item#']
finalList = {}
# f = open(".\Descritpion.csv","a")
for value in values:
    try:
        driver.get(f'https://www.shopcwi.com/searchresults.aspx?SearchTerm={value}')
        driver.maximize_window()
        xpath = '//*[@id="pagetabcontent_panel"]'
    except:
        print(value,", NULL")
    try:
        description = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        description = driver.find_element(By.XPATH,xpath).get_attribute("innerHTML")
        
        print(value,",",description)
        sheet1.write(rownum,0,value)
        sheet1.write(rownum,1,description)
        rownum+=1
        # finalList[value] = description
        # f.write(f'"{value}" , "{description}"\n')
    except:
        continue
# f.close()
wb.save('Description.xls')
driver.quit()