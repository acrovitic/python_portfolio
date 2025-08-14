import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

f = 'path/to/sheet_of_document_names.xlsx' # columns: filename, study#, meetingname
document_site = 'https://www.medresearchagency.gov/document_library/'
file_end_location_main = 'corporate/shared/drive/path/to/target_folder/'
hold_folder = "corporate/shared/drive/file_holding_folder/"

df=pd.read_excel(f)
df['download_url'] = f"{document_site}{df['PROTOCOL']}/{df['MEETING']}/{df['Name']}"
df['subfolder'] = f"{df['PROTOCOL']}/{df['MEETING']}"
df['old_file_path'] = hold_folder+df['Name']
df['new_file_path'] = f"{file_end_location_main}{df['subfolder']}/{df['Name']}"
downloaded_files = df.to_dict('records') # dictionary data structures allow for fast iteration
f_list = df['download_url'].tolist()

#initiate file download with selenium
options = webdriver.ChromeOptions()
tgt = hold_folder #download all files to one location to move as needed once selenium completes its task
profile = {"plugins.plugins_list": [
    {
        "enabled":False,
        "name":"Chrome PDF Viewer"
    }
],
    "download.default_directory" : tgt}
options.add_experimental_option("prefs",profile)
chromedriver = 'path/to/chromedriver.exe'
driver = webdriver.Chrome(chromedriver,options=options)
driver.get('website.com/file.pdf')
username = driver.find_element_by_id("txtUserName")
password = driver.find_element_by_id("txtPassword")
username.send_keys(username)
password.send_keys(password,Keys.RETURN)
username = driver.find_element_by_id("txtUserName")
password = driver.find_element_by_id("txtPassword")
username.send_keys(username)
password.send_keys(password,Keys.RETURN)
for i in f_list:
    page=driver.get(str(i))

not_moved=[]
for d in downloaded_files:
    if not os.path.exists(d['new_file_path']):
        try:
            os.makedirs(d['new_file_path'])
            os.rename(d['old_file_path'],d['new_file_path'])
        except:
            not_moved.append(d['new_file_path'])
            continue
    else:
        try:
            os.rename(tgt+'/'+d['Name'],d['new_file_path'])
        except:
            not_moved.append(d['new_file_path'])
            continue
