import traceback
import pandas as pd
import json
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

from .helpers.classes import *
from .data_cleaning.data_prep import *
pd.set_option('display.max_colwidth', -1)
pd.set_option('display.max_rows', 500)
pd.options.mode.chained_assignment = None

if __name__ == '__main__':
    # metadata of documents to upload
    target_file='documents_to_upload.xlsx'
    datapath = 'path/to/datadump/'
    upload_posting_folder_test = 'path/to/test/'
    upload_posting_folder = 'path/to/posting_uploading/'

    mtg_dict_narrowed, docs_final_dict, cpm_dict, df = get_prepped_data(
        target_file, datapath, upload_posting_folder_test, upload_posting_folder
    )

    #clean and bundle document data into dictionaries and append to a list of dictionaries
    to_upload=[]
    for doc in docs_final_dict:
        try:
            o = document(doc)
            o.get_recipients(cpm_dict)
            o.get_uploading_name()
            o.associate_protocol()
            o.get_mtg_data(mtg_dict_narrowed)
            o.clean_uploading_name()
            o.get_mtg_button_type()
            o.get_location()
            o.get_clean_location()
            o.get_destination()
            o.get_posting_association()
            to_upload.append(vars(o))
        except Exception as e:
            print(e)
            print(doc)
            traceback.print_exc()
            break

    # selenium based upload for documents to new web-based document repository and
    # append dicts of successfully uploaded docs to uploaded_docs list
    i = 0
    uploaded_docs = []
    chromedriver = 'path/to/chromedriver.exe'
    opt = Options()
    opt.add_argument("--start-maximized")
    driver = webdriver.Chrome(chromedriver, options=opt)
    driver.get('https://document_uploading_site.com')
    username = driver.find_element_by_id("txtUserName")
    password = driver.find_element_by_id("txtPassword")
    username.send_keys(username)
    password.send_keys(password,Keys.RETURN)
    for d in to_upload:
        file_upload = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, "upload")))
        select_prot = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ProtocolID']")))
        file_upload.send_keys(d["clean_location"])
        time.sleep(1)
        select_prot.send_keys(d["ProtocolName"])
        time.sleep(1)
        select_mtg = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "MeetingNameID"))
        )
        select_postmeeting_button = driver.find_element_by_css_selector("input[type='radio'][value='Post-Meeting']")
        select_meeting_button = WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"input[type='radio'][value='Meeting']"))
        )
        select_access = Select(driver.find_element_by_id("DocumentAccess"))

        select_doctype = Select(driver.find_element_by_id("DocumentTypeID"))
        select_doctype_list=[option.text for option in select_doctype.options]

        select_status = Select(driver.find_element_by_id("DocumentStatusesID"))
        select_version = driver.find_element_by_id("OfficialVersion")
        select_mtg.send_keys(d['mtg_name'])
        if d['mtg_name']!='select':
            if d['mtg_button_type']=='Post-Meeting':
                select_postmeeting_button.click()
            else:
                select_meeting_button.click()
        select_access.select_by_visible_text(d["access"])
        if d['DOC_TYPE'] in select_doctype_list:
            select_doctype.select_by_visible_text(d["DOC_TYPE"])
        else:
            select_doctype.send_keys(d["DOC_TYPE"])
        select_status.select_by_visible_text("Final")
        select_version.send_keys(d["version"])
        next_button=WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,".blue-button.nxtButton.col-md-12"))
        )
        next_button.click()
        upload_button = WebDriverWait(driver,10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='sample']"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", upload_button)
        upload_button.click()
        time.sleep(1)
        upload_ok = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "DocumentUploadOK"))
        )
        upload_ok.click()
        time.sleep(1)
        uploaded_docs.append(d)
        i+=1

    # output excel report of docs_uploaded
    out_path = "path/to/outputreport_folder/batch_completed.xlsx"
    writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
    workbook = writer.book
    df.to_excel(writer, sheet_name='Sheet1',index=False)
    worksheet = writer.sheets['Sheet1']
    writer.save()
