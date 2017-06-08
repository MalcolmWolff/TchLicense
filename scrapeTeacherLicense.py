##########################################
# Title:	TEACHER LICENSE SCRAPER - June 8, 2017
# Author:	Malcolm Wolff
##########################################

# Input:	'path_to_file.xlsx' is expected to be an
# 			xlsx file with the first column a list of teacher
# 			first names and the second column a list of teacher
# 			last names.
#
# Output:	'LicenseTable.xlsx' outputs an excel file with the following
#			information, in order of columns: License Number, First Name,
#			Middle Name, Last Name, plus a trash column that originally contained
#			a link.
#
# IMPORTANT:
#	THE CAPTCHA SYSTEM IS IFFY. IT MAY ONLY ASK YOU TO CLICK, WHICH THIS PROGRAM DOES
#	AUTOMATICALLY. IF IT ASKS YOU TO IDENTIFY PICTURES/ETC. YOU MUST DO THIS MANUALLY 
#	ONLY ONCE. CHANGE THE PARAMETER 'waitTime' BELOW TO ADJUST HOW MUCH TIME YOU NEED
#	TO FILL THIS OUT.
waitTime = 30




from pynput.mouse import Button, Controller
import os
from selenium import webdriver
from pyvirtualdisplay import Display
import time
from openpyxl import Workbook
from pandas import *

#Retrieve Name List
xls = ExcelFile('path_to_file.xlsx')
df = xls.parse(xls.sheet_names[0])
first_names = df.ix[:,0].to_dict()
last_names = df.ix[:,1].to_dict()

#Start Excel Document
wb = Workbook()
ws = wb.active

#Start Webdriver
display = Display(visible=0,size=(800,600))
display.start()

driver = webdriver.Chrome()
driver.get("https://tdoe.tncompass.org/Public/Search")

#Wait for Captcha to Load
time.sleep(5)



mouse = Controller()

#Read pointer position
#print('The current pointer position is {0}'.format(
#	mouse.position))

#Set pointer position
mouse.position = (168.3671875, 774.80859375)

#wait for mouse to move
time.sleep(waitTime)

#Press and release
mouse.press(Button.left)
mouse.release(Button.left)

mouse.press(Button.left)
mouse.release(Button.left)

mouse.press(Button.left)
mouse.release(Button.left)

#Get first name element
fname = driver.find_element_by_xpath('//*[@id="content"]/div/div/div[1]/div/div[1]/input')
lname = driver.find_element_by_xpath('//*[@id="content"]/div/div/div[1]/div/div[2]/input')

i = 0
for x in range(0,len(first_names)):
	fname.send_keys(first_names[x])
	lname.send_keys(last_names[x])

	time.sleep(5)

	#Get table
	table = driver.find_elements_by_xpath('//*[@id="educators-table"]/table/tbody/tr')
	for r in table:
		i = i+1
		j=0
		for td in r.find_elements_by_xpath('.//td'):
			j = j + 1
			ws.cell(row = i,column = j,value=td.text)

	fname.clear()
	lname.clear()

wb.save('LicenseTable.xlsx')
