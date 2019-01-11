# modules
import win32com
import win32com.client as win32
import os
import numpy as np
import pandas as pd
import re
import glob
import time
from dateutil.parser import parse
from collections import Counter
from shutil import copyfile
import docx2txt
import operator
import itertools
import datetime as dt
from bs4 import BeautifulSoup
import urllib.request
from dateutil.parser import parse
import requests
import shutil
from custom_classes import *
fro docprep_support_functions import *
from general_support_functions import *
from email_templates import *
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

df=pd.read_excel('path/to/prot_branch_fulltitle_spotfire.xlsx')
branch_dict=df.to_dict('records')

# initialize outlook folders
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
main_folder = outlook.Folders[2]
inbox=main_folder.Folders["Inbox"]
reports_folder = main_folder.Folders["Inbox"].Folders["Reports"]

outputs = []
for email in list(main_folder.Folders["Inbox"].Items):
    if "New Protocol Version Available:" in email.Subject:
        file_urls = re.findall('<(.*?)>',email.Body)
        for url in file_urls:
            if "http" in url:
                o = protocol_doc(url)
                o.get_new_doc_attributes()
                outputs.append(vars(o))

# group docs of same protocol number, filter and dedup for easier posting references
group_by(outputs,'protocol','new_file_path')
group_by(outputs,'protocol','path')
prefinal_d=deduplicate_filtered_keys(
    filter_keys(
        outputs,['protocol','title','path_','new_file_path_','title','version','date','team']
    )
)
deduped_final_d=remove_dupes(prefinal_d, "protocol").values()

final_d=[]
for d in deduped_final_d:
    final_d.append(d)

ready_to_post = []
for d in final_d:
    pkg = doc_package(d)
    pkg.rename_downloaded_files()
    pkg.set_email_msg_part()
    pkg.send_icon_new_docs()
    pkg.move_docs_on_mdrive()
    pkg.get_posting_metadata(branch_dict)
    ready_to_post.append(vars(pkg))
    
# upload Protocols to DL and CMS
chromedriver = 'path/to/chromedriver.exe'
opt = Options() 
opt.add_argument("--start-maximized") 
driver=webdriver.Chrome(chromedriver, options=opt) 
driver.get('www.website1.com')
username = driver.find_element_by_id("txtUsername")
password = driver.find_element_by_id("txtPassword")
username.send_keys(username)
password.send_keys(password,Keys.RETURN)
driver.get('www.website2.com')
username = driver.find_element_by_id("txtUserName")
password = driver.find_element_by_id("txtPassword")
username.send_keys(username)
password.send_keys(password,Keys.RETURN)
for d in ready_to_post:
    driver.get('www.website1.com')
    old_dl_button=WebDriverWait(driver,10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR,"#zz11_RootAspMenu > li > ul > li:nth-child(2) > a"))
    )
    old_dl_button.click()
    window_before = driver.window_handles[0]
    window_now=driver.window_handles[1]
    driver.switch_to.window(window_before)
    driver.close()
#     another_window = list(set(driver.window_handles) - {driver.current_window_handle})[0]#refocuses on new window
    driver.switch_to.window(window_now)#refocuses on new window
    old_dl_upload=WebDriverWait(driver,15).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR,"#zz12_RootAspMenu > li:nth-child(2) > ul > li > a > span > span"))
    )
    old_dl_upload.click()

    old_dl_add_file=WebDriverWait(driver,10).until(
        EC.element_to_be_clickable((By.XPATH,'//*[@id="idHomePageNewDocument"]'))
    )
    old_dl_add_file.click()
    
    iframe = driver.find_element_by_xpath('//iframe[contains(@id, "Frame")]') # works to switch over
    driver.switch_to.frame(iframe) 
    
    old_dl_choose_file=WebDriverWait(driver,10).until(
        EC.presence_of_element_located((By.NAME,"ctl00$PlaceHolderMain$UploadDocumentSection$ctl03$InputFile"))
    )
    old_dl_choose_file.send_keys(d['posted_path'])
    driver.find_element_by_name("ctl00$PlaceHolderMain$UploadDocumentSection$ctl03$OverwriteSingle").click()
    driver.find_element_by_css_selector("#ctl00_PlaceHolderMain_ctl03_RptControls_btnOK").click()
    
    old_content_type=WebDriverWait(driver,10).until(
        EC.presence_of_element_located((By.NAME,"ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$ctl01$ctl00$ContentTypeChoice"))
    )
    old_content_type.send_keys(d['old_dl_doctype'])
    WebDriverWait(driver,10).until(
        EC.text_to_be_present_in_element((By.NAME,"ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$ctl01$ctl00$ContentTypeChoice"), d['old_dl_doctype'])
    )
    # protocol number
    driver.find_element_by_name("ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$ctl02$ctl00$ctl00$ctl02$ctl00$ctl00$ctl04$ctl00$Lookup").send_keys(d['protocol'])
    # doc group
    driver.find_element_by_name("ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$ctl02$ctl00$ctl00$ctl04$ctl00$ctl00$ctl04$ctl00$DropDownChoice").send_keys('SOCS documents')
    # doc type
    driver.find_element_by_name("ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$ctl02$ctl00$ctl00$ctl05$ctl00$ctl00$ctl04$ctl00$DropDownChoice").send_keys(d['old_dl_doctype'])
    # doc status (i.e. access)
    driver.find_element_by_name("ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$ctl02$ctl00$ctl00$ctl09$ctl00$ctl00$ctl04$ctl00$DropDownChoice").send_keys(d['access'])
    # DID #
    driver.find_element_by_name("ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$ctl02$ctl00$ctl00$ctl11$ctl00$ctl00$ctl04$ctl00$ctl00$TextField").send_keys("1")
    # other access dropdown 
    driver.find_element_by_name("ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$ctl02$ctl00$ctl00$ctl14$ctl00$ctl00$ctl04$ctl00$DropDownChoice").send_keys(d['access'])
    # branch
    driver.find_element_by_id('ctl00_m_g_8909dea4_97f6_4339_a9b0_3697b55f26ad_ctl00_ctl02_ctl00_ctl02_ctl00_ctl00_ctl13_ctl00_ctl00_ctl04_ctl00_ctl02editableRegion').send_keys(d['Branch'])
     #smc meeting
    driver.find_element_by_xpath('//*[@id="ctl00_m_g_8909dea4_97f6_4339_a9b0_3697b55f26ad_ctl00_ctl02_ctl00_ctl02_ctl00_ctl00_ctl15_ctl00_ctl00_ctl04_ctl00_ctl00_TextField"]').click()
    driver.find_element_by_xpath('//*[@id="ctl00_m_g_8909dea4_97f6_4339_a9b0_3697b55f26ad_ctl00_ctl02_ctl00_ctl02_ctl00_ctl00_ctl15_ctl00_ctl00_ctl04_ctl00_ctl00_TextField"]').send_keys(d['old_dl_mtg'])
    # ok button
    driver.find_element_by_name("ctl00$m$g_8909dea4_97f6_4339_a9b0_3697b55f26ad$ctl00$ctl02$ctl00$toolBarTbl$RightRptControls$ctl00$ctl00$diidIOSaveItem").click()
    time.sleep(7)
    driver.get('www.website2.com')
    file_upload = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "upload"))
    )
    file_upload.send_keys(d["posted_path"])
    select_prot = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ProtocolID']"))
    )
    select_prot.send_keys(d["protocol"])
    time.sleep(1)
    select_access = Select(driver.find_element_by_id("DocumentAccess"))
    select_doctype = Select(driver.find_element_by_id("DocumentTypeID"))
    select_status = Select(driver.find_element_by_id("DocumentStatusesID"))
    select_version = driver.find_element_by_id("OfficialVersion")
    select_access.select_by_visible_text(d["access"])
    select_doctype.select_by_visible_text(d["doc_type"])
    select_status.select_by_visible_text(d["status"])
    select_version.send_keys(d["version"].replace("v",""))
    driver.find_element_by_css_selector(".cms-blue-button.nxtButton.col-md-12").click()
    time.sleep(1)
    upload_button = WebDriverWait(driver,20).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='sample']"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", upload_button)
    upload_button.click()
    time.sleep(1)
    upload_ok = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "DocumentUploadOK"))
    )
    upload_ok.click()
    time.sleep(1)
