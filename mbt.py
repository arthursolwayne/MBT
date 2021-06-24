"""

Author: Arthur Wayne
Date: June 6 2021
Description: CME Micro Bitcoin (MBT) Futures two-factor trading algorithm.

"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

import csv
import smtplib
from datetime import datetime
import pyautogui
import pytesseract
from desktopmagic.screengrab_win32 import getRectAsImage
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

def grabData (openTrade, size):
	# time
	now = datetime.now()
	dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
	# volume factor
	v = pytesseract.image_to_string(getRectAsImage((11, 208, 132, 362)))
	p1SD = "T" in v[v.find("P1")+3:v.find("P1")+7]
	p2SD = "T" in v[v.find("P2")+3:v.find("P2")+7]
	n1SD = "T" in v[v.find("N1")+3:v.find("N1")+7]
	n2SD = "T" in v[v.find("N2")+3:v.find("N2")+7]
	bounded = not (p1SD == True or n1SD == True)
	clo = pytesseract.image_to_string(getRectAsImage((11, 363, 132, 398)))
	close = int(clo[clo.find(":")+2:clo.find(".")])
	vw = pytesseract.image_to_string(getRectAsImage((11, 402, 132, 437)))
	vwap = int(vw[vw.find(":")+2:vw.find(".")])
	# m1 factor
	m1 = pytesseract.image_to_string(getRectAsImage((11, 493, 132, 644)))
	m1UF = "T" in m1[m1.find("UF")+3:m1.find("UF")+7]
	m1DF = "T" in m1[m1.find("DF")+3:m1.find("DF")+7]
	m1_1SD = "T" in m1[m1.find("1SD")+4:m1.find("1SD")+8]
	m1_2SD = "T" in m1[m1.find("2SD")+4:m1.find("2SD")+8]
	# m5 factor
	m5 = pytesseract.image_to_string(getRectAsImage((11, 830, 132, 983)))
	m5UF = "T" in m5[m5.find("UF")+3:m5.find("UF")+7]
	m5DF = "T" in m5[m5.find("DF")+3:m5.find("DF")+7]
	m5_1SD = "T" in m5[m5.find("1SD")+3:m5.find("1SD")+7]
	m5_2SD = "T" in m5[m5.find("2SD")+3:m5.find("2SD")+7]
	return [openTrade, size, dt_string, close, vwap, bounded, p1SD, p2SD, n1SD, n2SD, m1UF, m1DF, m1_1SD, m1_2SD, m5UF, m5DF, m5_1SD, m5_2SD]

def makeTrade (ss):
	chrome_options = Options()
	chrome_options.add_argument("--window-size=800,1055")
	driver = webdriver.Chrome(chrome_options=chrome_options)
	driver.set_window_position(185, 0, windowHandle='current')
	driver.get('https://www.cmegroup.com/futures_challenge/simulator')
	user = driver.find_element_by_id('userName').send_keys('')
	pword = driver.find_element_by_id('password').send_keys('')
	login = driver.find_element_by_class_name('btn--primary').click()
	trade = driver.find_element_by_xpath('//*[@id="symbol_MBT"]/div/div/a').click()
	otype = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[1]/div[2]/div[1]/div/div/button/a').click()
	mkt = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[1]/div[2]/div[1]/div/div/ul/li[1]/a').click()
	qty = driver.find_element_by_id("orderQuantity").send_keys("10")
	# if p2SD sell to open
	if ss[7] == True:
		side = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[1]/div[3]/div[2]/div/span/label[2]').click()
		ss[1] = -10
	# if n2SD buy to open
	if ss[9] == True:
		side = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[1]/div[3]/div[2]/div/span/label[1]').click()
		ss[1] = 10
	submit = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[2]/div[4]/div[2]/input[1]').click()
	ok = driver.find_element_by_xpath('//button[text()="OK"]').click()
	ok2 = driver.find_element_by_xpath('/html/body/div[9]/div/div[2]/div[2]/div/div[2]/input[1]').click()
	ss[0] = True
	print("Position Opened at "+str(ss[3])+". Size: "+ str(ss[1]))
	return ss

def closeTrade (ss):
	chrome_options = Options()
	chrome_options.add_argument("--window-size=800,1055")
	driver = webdriver.Chrome(chrome_options=chrome_options)
	driver.set_window_position(185, 0, windowHandle='current')
	driver.get('https://www.cmegroup.com/futures_challenge/simulator')
	user = driver.find_element_by_id('userName').send_keys('')
	pword = driver.find_element_by_id('password').send_keys('')
	login = driver.find_element_by_class_name('btn--primary').click()
	trade = driver.find_element_by_xpath('//*[@id="symbol_MBT"]/div/div/a').click()
	otype = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[1]/div[2]/div[1]/div/div/button/a').click()
	mkt = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[1]/div[2]/div[1]/div/div/ul/li[1]/a').click()
	qty = driver.find_element_by_id("orderQuantity").send_keys("10")
	# buy to close
	if ss[1] < 0:
		side = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[1]/div[3]/div[2]/div/span/label[1]').click()
		ss[1] = 0
	# sell to close
	if ss[1] > 0:
		side = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[1]/div[3]/div[2]/div/span/label[2]').click()
		ss[1] = 0
	submit = driver.find_element_by_xpath('//*[@id="order-ticket"]/div[2]/div[4]/div[2]/input[1]').click()
	ok = driver.find_element_by_xpath('//button[text()="OK"]').click()
	ok2 = driver.find_element_by_xpath('/html/body/div[9]/div/div[2]/div[2]/div/div[2]/input[1]').click()
	ss[0] = False
	print("Position Closed at "+str(ss[3])+". Size: "+ str(ss[1]))
	return ss

def sendOpenEmail(open):
	mail = smtplib.SMTP('smtp.gmail.com',587)
	mail.ehlo()
	mail.starttls()
	mail.login('','')
	if open == True:
		message = "Trade Opened"
	else:
		message = "Trade Closed"
	mail.sendmail('','', message)
	mail.close()


openTrade = False
size = 0
with open('mbtData.csv', 'a', newline = '') as f:
	thewriter = csv.writer(f)
	print("Welcome back. Let's win this competition!")
	print('Working...press ctrl + c to exit.')
	thewriter.writerow(['OpenTrade', 'Size', 'Time', 'Close', 'VWAP', 'Bounded', 'P1SD', 'P2SD', 'N1SD', 'N2SD', 'm1UF', 'm1DF', 'm1_1SD', 'm1_2SD', 'm5UF', 'm5DF', 'm5_1SD', 'm5_2SD'])
	while(1):
		try:
			ss = grabData(openTrade, size)
			if ss[0] == False and (ss[7] == True or ss[9] == True):
				ss = makeTrade(ss)
				openTrade = ss[0]
				size = ss[1]
				thewriter.writerow(ss)
				sendOpenEmail(True)
			elif ss[0] == True and ((ss[1]<0 and ss[3]<=ss[4]) or (ss[1]>0 and ss[3]>=ss[4])):
				ss = closeTrade(ss)
				openTrade = ss[0]
				thewriter.writerow(ss)
				sendOpenEmail(False)
			elif datetime.now().second == 0 or (ss[0] == False and ((ss[1]<0 and ss[3]<=ss[4]) or (ss[1]>0 and ss[3]>=ss[4]))):
				if openTrade == True:
					entry = 0
					pnl = (ss[3]-entry)*ss[1]
					print(pnl)
				thewriter.writerow(ss)
				if datetime.now().minute == 0:
					print("New Hour. Time: "+ss[2])
		except KeyboardInterrupt:
			break