"""

This file pulls 3 years of learnpro assessment attempts in 6 month increments, then saves them in a folder on the
W drive where they can be used to feed the learnpro named list creator.

"""

from selenium import webdriver
import time
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
import os
import pandas as pd
import configparser
from pandas.tseries.offsets import MonthEnd
from dateutil.relativedelta import *

# get current date
current = pd.to_datetime(pd.Timestamp('now').strftime("%Y-%m-%d 00:00:00"))


today = datetime.today()
targetdate = pd.to_datetime(input("What's the target month? (format = MM/YYYY)")) + MonthEnd(1)
folder = 'W:/Learnpro/Data/' + str(today.strftime('%Y%m%d')) + '-auto/'
if not os.path.isdir(folder):
    os.mkdir(folder)
filename = r"W:\MicroStrategy\Data\Bank\Assessment_Attempts_-_All_Assessments.xls"




users = r"W:\MicroStrategy\Data\Bank\Users_-_Current_Users.xls"
if os.path.exists(filename):
    os.remove(filename)
if os.path.exists(users):
    os.remove(users)

config = configparser.ConfigParser()
config.read(r'W:\\Python\Danny\Learnpro Automated Extract\learnpro.ini')
print(config.get('learnpro', 'uname'))

chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": "W:\MicroStrategy\Data\Bank"
         }
chromeOptions.add_experimental_option("prefs", prefs)
#chromeOptions.add_argument("--headless")
browser = webdriver.Chrome(executable_path=r"W:\Danny\\Chrome Webdriver\\chromedriver.exe",
                           options=chromeOptions)

typelist ={'First Name':str, 'Last Name':str, 'Email / Username':str,
         'ID Number':str, 'Course':'category','Module':'category', 'Score':int,
         'Out Of':int, 'Duration (s)':str, 'Passed':'category',
         'Hospital':'category',
         'Ward':'category','Job Family':'category', 'Sub Family':'category',
         'Role':'category','Directorate':'category', 'Subdirectorate':'category'}

def login():
    browser.get('https://nhs.learnprouk.com/lms/login.aspx')
    time.sleep(2)
    username = browser.find_element_by_id('ContentPlaceHolder1_TextBoxUsername')
    password = browser.find_element_by_id('ContentPlaceHolder1_TextBoxPassword')
    username.send_keys(config.get('learnpro', 'uname'))
    password.send_keys(config.get('learnpro', 'pword'))
    time.sleep(1)
    browser.find_element_by_id('ContentPlaceHolder1_ButtonLogin').click()


def useraccounts():
    browser.get('https://nhs.learnprouk.com/lms/user_Level/reportcreator.aspx')
    time.sleep(1.5)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_LabelCategoryName_18').click()
    time.sleep(1.5)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_DataListReports_18_LinkButtonReport_0').click()


def reportcreator():
    browser.get('https://nhs.learnprouk.com/lms/user_Level/reportcreator.aspx')
    time.sleep(1.5)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_LabelCategoryName_3').click()
    time.sleep(1.5)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_DataListReports_3_LinkButtonReport_0').click()


year1_1 = targetdate - relativedelta(months=6)
year1_2 = year1_1 - relativedelta(months=6)
year2_1 = year1_2 - relativedelta(months=6)
year2_2 = year2_1 - relativedelta(months=6)
year3_1 = year2_2 - relativedelta(months=6)
year3_2 = year3_1 - relativedelta(months=6)



def runthrough(date, enddate, filez):
    reportcreator()
    time.sleep(2)
    dateinput(date, enddate)
    time.sleep(3)
    get_through()
    counter = 0
    while not os.path.exists(filename):
        time.sleep(2)
        counter += 1
        print(counter)
        if counter == 30:
            print('timeout')
    time.sleep(2)
    os.rename(filename, folder+'LEARNPRO '+enddate.strftime('%Y%m%d')+"-"+date.strftime('%Y%m%d')+'.xls')

def dateinput(date, enddate):
    endtime = browser.find_element_by_id('ctl00_ContentPlaceHolder1_ReportCreator1_DatePickerFromDate_Picker1_picker')
    starttime = browser.find_element_by_id('ctl00_ContentPlaceHolder1_ReportCreator1_DatePickerToDate_Picker1_picker')
    starttime.clear()

    starttime.send_keys(Keys.RIGHT)
    time.sleep(1)
    starttime.send_keys(date.strftime('%m'))
    time.sleep(1)
    starttime.send_keys(Keys.LEFT)
    time.sleep(1)
    starttime.send_keys(date.strftime('%d'))
    time.sleep(1)
    starttime.send_keys(Keys.RIGHT)
    time.sleep(1)
    starttime.send_keys(Keys.RIGHT)
    time.sleep(1)
    starttime.send_keys(date.strftime('%y'))
    # input_element.send_keys(:arrow_down)
    endtime.clear()


    endtime.send_keys(Keys.RIGHT)
    time.sleep(1)
    endtime.send_keys(enddate.strftime('%m'))
    time.sleep(1)
    endtime.send_keys(Keys.LEFT)
    time.sleep(1)
    endtime.send_keys(enddate.strftime('%d'))

    time.sleep(1)
    endtime.send_keys(Keys.RIGHT)
    time.sleep(1)
    endtime.send_keys(Keys.RIGHT)
    time.sleep(1)
    endtime.send_keys(enddate.strftime('%y'))
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_ButtonConfirmDate').click()


def get_through():
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_ButtonConfirmLocationRole').click()
    time.sleep(2)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_ButtonConfirmDirectorate').click()
    time.sleep(2)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_HyperlinkCSVXLSViewer').click()


def get_through_disability():
    time.sleep(2)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_ButtonConfirmLocationRole').click()
    time.sleep(2)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_ButtonConfirmDirectorate').click()
    time.sleep(2)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_ButtonConfirmDisabledAccounts').click()
    time.sleep(2)
    browser.find_element_by_id('ContentPlaceHolder1_ReportCreator1_HyperlinkCSVXLSViewer').click()

    # print(learnpro.credentials)


login()
time.sleep(2)
useraccounts()
time.sleep(2)
get_through_disability()
counter = 0
while not os.path.exists(users):
    time.sleep(2)
    counter += 1
    print(counter)
    if counter == 30:
        print('timeout')
#userdf = pd.read_csv(users, skiprows=10, sep='\t')
#userdf.to_excel('W:/Learnpro/Data/Users/'+today.strftime('%Y%m%d')+" All Users.xlsx", engine= 'xlsxwriter', index=False)
os.rename(users, folder + 'users.xls')
runthrough(current, targetdate, 'year0.xlsx')
time.sleep(5)
runthrough(targetdate, year1_1, 'year1_1.xlsx')
time.sleep(5)
runthrough(year1_1, year1_2, 'year1_2.xlsx')
time.sleep(5)
runthrough(year1_2, year2_1, 'year2_1.xlsx')
time.sleep(5)
runthrough(year2_1, year2_2, 'year2_2.xlsx')
time.sleep(5)
runthrough(year2_2, year3_1, 'year3_1.xlsx')
time.sleep(5)
runthrough(year3_1, year3_2, 'year3_2.xlsx')


print("Finished")