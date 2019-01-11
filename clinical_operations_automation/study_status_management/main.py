import os
import re
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import glob
import datetime as dt
from dateutil.parser import parse
from collections import Counter
from shutil import copyfile
import operator
from functions import *
from email_template import *
import itertools
import shutil
import time
pd.set_option('display.max_colwidth', -1)
today = dt.datetime.now().strftime('%d%b%y')

# download current polling calendar
options = webdriver.ChromeOptions()
profile = {"plugins.plugins_list": [{"enabled":False, "name":"Chrome PDF Viewer"}],}
options.add_experimental_option("prefs",profile)
chromedriver = 'C:/Python Scripts/chromedriver/chromedriver.exe'
driver = webdriver.Chrome(chromedriver,options=options)
chromedriver = 'C:/Python Scripts/chromedriver/chromedriver.exe'
opt = Options() 
opt.add_argument("--start-maximized") 
driver=webdriver.Chrome(chromedriver, options=opt) 
driver.get('www.website.com')
username = driver.find_element_by_id("txtUserName")
password = driver.find_element_by_id("txtPassword")
username.send_keys(username)
password.send_keys(password,Keys.RETURN)
page = driver.get("www.website.com/reports/polling_calendar_report")
driver.find_element_by_id("ExecuteQueryButton").click()
time.sleep(1)
driver.find_element_by_xpath('//*[@id="ResultExportButtons"]/a[2]').click()
time.sleep(1)
driver.quit()

# get latest files, find items in main polcal missing from website
polcal_path = "path/to/exported/report"
all_polcals = glob.glob(polcal_path + 'Polling Calendar*')
latest_polcal = max(all_polcals, key=os.path.getctime)

main_polcal_path = "path/to/main/report"
all_main_polcals = glob.glob(main_polcal_path + 'PollingCalendar*')
latest_main_polcal = max(all_main_polcals, key=os.path.getctime)

ipath = "path/to/study/assignments/"
list_of_files = glob.glob(ipath + 'Staff Protocol Assignments*')
latest_file = max(list_of_files, key=os.path.getctime)
df2 = pd.read_excel(latest_file,skiprows=1)

# find items in main polling calendar that are missing from the CMS
df_main_polcal = pd.read_excel(latest_main_polcal,skiprows=1)
df_polcal = pd.read_excel(latest_polcal)
df_main_polcal = df_main_polcal[~df_main_polcal['Branch'].str.contains('Completed')]
df_polcal['Meeting Type'] = df_polcal['Meeting Type'].apply(lambda x: mtg_switch(x))
df_polcal['unique_key'] = df_polcal['Protocol Number'] + df_polcal['Meeting Type']
df_main_polcal['unique_key'] = df_main_polcal[' Protocol Number'] + df_main_polcal['Meeting Type']
df_missing = df_main_polcal[~df_main_polcal['unique_key'].isin(df_polcal['unique_key'])]
include = ['ORG','DRM','Ad Hoc']
dfm = df_missing[~((df_missing['Date & Time Scheduled']=='Pending') & 
           (df_missing['Meeting Type'].isin(include)))]
dfm = dfm[[' Protocol Number','Meeting Type','Polling Dates','Date & Time Scheduled']].copy()
df2 = df2[['Protocol #','Staff']].copy()
df2.columns = [' Protocol Number','Staff']
dfm = dfm.merge(df2,on=' Protocol Number',how='left')
dfm['email'] = dfm['TRI'].apply(lambda x: get_email_address(x))
dfm['name'] = dfm['TRI'].apply(lambda x: get_name(x))
required_updates = dfm.to_dict('records')
for d in required_updates:
    if 'polling' in d['Date & Time Scheduled'].lower():
        d['action'] = 'Add Protocol {p} {m} polling.'.format(p=d[' Protocol Number'],m=d['Meeting Type'])
    if '@' in d['Date & Time Scheduled'].lower():
        d['action'] = 'Add Protocol {p} meeting scheduled for {d}.'.format(p=d[' Protocol Number'],d=d['Date & Time Scheduled'])
    if d['Meeting Type'] == 'E-Rev':
        if 'pending' in d['Date & Time Scheduled'].lower():
            d['action'] = 'Add Protocol {p} E-Review started {d}.'.format(p=d[' Protocol Number'],d=d['Polling Dates'])
        if 'pending' not in d['Date & Time Scheduled'].lower():
            d['action'] = 'Add Protocol {p} E-Review completed {d}.'.format(p=d[' Protocol Number'],d=d['Date & Time Scheduled'])

# group and dedup
group_by(required_updates,'name','action')
prefinal_d=deduplicate_filtered_keys(
    filter_keys(
        required_updates,['email', 'name', 'action_']
    )
)
deduped_final_d=remove_dupes(prefinal_d, "email").values()
final_d=[]
for d in deduped_final_d:
    final_d.append(d)
for d in final_d:
    if len(d['action_']) > 1:
        d['actions'] = '<br>'.join(d['action_'])
    if len(d['action_']) == 1:
        d['actions'] = ''.join(d['action_'])

# write emails to team members who need to update their study's data on the website
for d in final_d:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = d['email']
    mail.Subject = 'Updates Required - ' + today
    mail.HTMLBody = email.format(**d)
    mail.Display(False)
