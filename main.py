from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import time, datetime
import pandas as pd
import numpy as np
from openpyxl import Workbook
from pytesseract import image_to_string
import cv2 as cv

excel = Workbook()
excel_sheet = excel.create_sheet('기업정보')
excel_sheet = excel.active
excel_sheet.append(
    ["기업체명", "영문기업명", "사업자번호", "법인(주민)번호", "대표자명", "종업원수", "설립형태", "설립일자", "기업형태", "기업규모", "전화번호", "팩스번호", "홈페이지",
     "이메일", "결산원", "기업공개일자", "도로명", "지번", "업종(10차)", "업종(9차)", "주요제품(상품)", "무역업허가번호", "소속그룹", "주채권기관", "당좌거래은행",
     "휴페업정보", "법인등기정보"])
driver = webdriver.Chrome(os.getcwd() + "\\chromedriver.exe")
wait = WebDriverWait(driver, 5)
main_page = driver.current_window_handle


def captcha():
    try:
        print("캡챠인증 진행")
        image = driver.find_element_by_xpath('//*[@id="slythgomi"]')
        image = image.screenshot_as_png
        with open(os.getcwd() + "\\captcha.png", "wb") as file:
            file.write(image)
        image = cv.imread(os.getcwd() + "\\captcha.png")
        captcha_text = image_to_string(image)
        captcha_text.replace(" ", "")
        print(captcha_text)
        driver.find_element_by_xpath('//*[@id="certChar"]').send_keys(captcha_text)
        try:
            wait.until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except:
            wait.until(EC.alert_is_present())
            driver.switch_to.alert.accept()
    except:
        print("캡챠인증 진행")
        image = driver.find_element_by_xpath('//*[@id="slythgomi"]')
        image = image.screenshot_as_png
        with open(os.getcwd() + "\\captcha.png", "wb") as file:
            file.write(image)
        image = cv.imread(os.getcwd() + "\\captcha.png")
        captcha_text = image_to_string(image)
        captcha_text.replace(" ", "")
        print(captcha_text)
        driver.find_element_by_xpath('//*[@id="certChar"]').send_keys(captcha_text)
        try:
            wait.until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except:
            wait.until(EC.alert_is_present())
            driver.switch_to.alert.accept()


url = 'http://www.cretop.com/'
driver.get(url)
print("로그인중")
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="in_id"]'))).send_keys('k0230062')
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="in_pw"]'))).send_keys('kedkorea!23')
driver.find_element_by_xpath('//*[@id="loginBtn1"]').click()
print("로그인 완료")
print("스마트서치")
q = 1
k = 1
count = 0

while True:
    for q in range(1, 51):
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="CMSRC04R0"]'))).click()
        except:
            captcha()
            q -= 1
            continue
        print('page : ', count)
        print('count : ', count * 50 + q)
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="enpSze"]/li[1]'))).click()
        except TimeoutException:
            captcha()
            q -= 1
            continue

        while True:
            try:
                wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="selPage"]'))).click()
                break
            except:
                pass

        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="selPage"]/option[4]'))).click()
        except:
            captcha()
            q -= 1
            continue
        try:
            driver.find_element_by_xpath('//*[@id="CMSRC04S0DIV"]/div[1]/div/input').click()
        except:
            captcha()
            q -= 1
            continue
        for i in range(0, count):
            try:
                while True:
                    try:
                        wait.until(
                            EC.visibility_of_element_located(
                                (By.XPATH,
                                 '//*[@id="srchListDiv"]/div/div/div[2]/a[3]'))).click()
                        break
                    except:
                        pass
            except TimeoutException:
                captcha()
                q -= 1
                continue
        try:
            while True:
                try:
                    wait.until(EC.element_to_be_clickable(
                        (By.XPATH,
                         '//*[@id="srchListDiv"]/div/div/div[1]/table/tbody/tr[' + str(q) + ']/td[1]/a'))).click()
                    break
                except:
                    pass
        except:
            captcha()
            q -= 1
            continue

        try:
            while True:
                try:
                    wait.until(EC.visibility_of_element_located(
                        (By.XPATH,
                         '//*[@id="srchListDiv"]/div/div/div[1]/table/tbody/tr[' + str(
                             q) + ']/td[1]/div/div/div/ul/li[3]/a'))).click()
                    break
                except:
                    pass
        except TimeoutException:
            captcha()
            q -= 1
            continue
        try:
            while True:
                try:
                    print(driver.find_element_by_xpath('//*[@id="CMCOM13S1DIV"]/div/div/div/h4').text)
                except:
                    pass
                finally:
                    break
        except TimeoutException:
            captcha()
            q -= 1
            continue
        tmp_list = []
        for i in range(1, 17):
            for j in range(1, 3):
                if i in (9, 10, 11, 12, 13):
                    if j == 1:
                        try:
                            tmp_list.append(driver.find_element_by_xpath(
                                '//*[@id="frm"]/div[3]/table/tbody/tr[' + str(i) + ']/td[1]').text)
                        except:
                            captcha()
                            q -= 1
                            continue
                    else:
                        pass
                else:
                    try:
                        tmp_list.append(driver.find_element_by_xpath(
                            '//*[@id="frm"]/div[3]/table/tbody/tr[' + str(i) + ']/td[' + str(j) + ']').text)
                    except:
                        captcha()
                        q -= 1
                        continue
        excel_sheet.append(tmp_list)
        driver.execute_script("history.back();")
        try:
            while True:
                try:
                    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="enpSze"]/li[1]'))).click()
                except:
                    pass
                finally:
                    break
        except TimeoutException:
            captcha()
            q -= 1
            continue
        print('Next')
        excel.save(filename='Data.xlsx')
    count += 1
