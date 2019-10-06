import xlwt
from selenium import webdriver
from bs4 import BeautifulSoup
import requests
import pandas as pd
import xlwt
import win32com.client

page_num = int(input("마지막 페이지 번호를 입력해 주세요 : "))

address = 'https://minwon.nhis.or.kr/wbm/kb/retrieveHomeOxyProductList.xx'

driver = webdriver.Chrome('/Users/kikik/Downloads/chromedriver_win32/chromedriver')
driver.implicitly_wait(3)
driver.get(address)
# 아이디/비밀번호를 입력해준다.

#html = driver.page_source
#soup = BeautifulSoup(html, 'html.parser')
#print(soup)
# driver.find_element_by_partial_link_text("2").click()

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

wb = excel.Workbooks.Add()

ws = wb.Worksheets("Sheet1")

for num in range(2, page_num+2):
    if num%10 == 1:
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        num3 = 1
        num2 = 1
        for link2 in soup.find_all('td', {'class': 'cen'}):
            ws.Cells(num3 + (num - 2) * 10, num2).Value = link2.text.strip()
            num2 = num2 + 1
            if num2 % 8 == 0:
                num2 = 1
                num3 = num3 + 1

        element = driver.find_element_by_xpath('//a[img/@src="/static/images/egovframework/cmmn/btn_page_next1.gif"]')
        driver.execute_script("arguments[0].click();", element)
    else:
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        #for link2 in soup.find_all('td', {'class': 'cen'}):
        #    print(link2.text.strip())
        numstr = str(num)
        element = driver.find_element_by_partial_link_text(numstr)
        driver.execute_script("arguments[0].click();", element)
        num3 = 1
        num2 = 1
        for link2 in soup.find_all('td', {'class':'cen'}):
            ws.Cells(num3+(num-2)*10, num2).Value = link2.text.strip()
            num2 = num2+1
            if num2%8 == 0:
                num2 = 1
                num3 = num3+1
        #html = driver.page_source
        #soup = BeautifulSoup(html, 'html.parser')