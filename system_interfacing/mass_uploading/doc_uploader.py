# modules
import os
import re
import traceback
import pandas as pd
import numpy as np
import json
import time
from dateutil.parser import parse
import difflib
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
pd.set_option('display.max_colwidth', -1)
pd.set_option('display.max_rows', 500)
pd.options.mode.chained_assignment = None

#functions
def get_dates(string):
    string=str(string)
    month=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    match=re.search('(\d{1,2}\/\d{1,2}\/\d{2,4})|(\d{1,2}\ +\w{3,9}\ +\d{2,4})|(\d{1,2}\w{3}\d{2})|(\w{3,9}\ +\d{1,2},\s+\d{4})',string)
    mtgtypes=['data review','org','ad hoc','ad-hoc','e-rev','e rev','elec','drm','follow-up','follow up','meeting']
    if not any(i in string.lower() for i in mtgtypes):
        return np.nan
    else:
        if match:
            date=match[0]
            if len(date)==7 and re.match('(\d{1,2}\w{3}\d{2})',date):
                if any(i in date.lower() for i in month):
                    return parse(date).strftime('%m/%d/%Y')
                else:
                    return np.nan
            else:
                return parse(date).strftime('%m/%d/%Y')
        else:
            return np.nan
def date_matcher(string):
    string=str(string).replace("\ISM","")
    if 'H5N1' in string:
        string=string.replace('ABC-','').replace('-EFG','')
        return string
    if ":" in string:
        string_date_parts=string.rsplit('-',1)
        string_date_parts[1]=parse(''.join(string_date_parts[1].rsplit(' ',1)[:1])).strftime('%m/%d/%Y')
        return ' '.join(string_date_parts)
    else:
        return string.replace("DSMB-","DSMB 0")
def get_access(string):
    string=string.lower()
    if "close" in string or "unblind" in string:
        return "Closed"
    else:
        return "Open"    
def get_version(file_name):
    match1 = re.search("v(\d{1,2}\.\d{1})",file_name)
    match2 = re.search("(\d{1,2}\.\d{1})",file_name)
    if not match1 and not match2:
        return "1.0" # for batch 2 ONLY. change later
    elif not match1:
        return match2[1]
    else:
        return match1[1] 
def filter_keys(list_of_dictionaries,needed_keys_list):
    dict1 = []
    for d in list_of_dictionaries:
        filtered_d = dict((k, d[k]) for k in needed_keys_list if k in d)
        dict1.append(filtered_d)
    return dict1
def deduplicate_filtered_keys(filtered_list_of_dictionaries):
    dict2 = []
    for d in filtered_list_of_dictionaries:
        if d not in dict2:
            dict2.append(d)
    return dict2

#classes
#custom class to add sleep action to Selenium action chains
class Actions(ActionChains):
    def wait(self, time_s: float):
        self._actions.append(lambda: time.sleep(time_s))
        return self

# uploading class to further clean data and associate related documents together for later posting
class document(object):
    protnum_associations={'old assoc':'new assoc'}
    old_path="path/to/files/old"
    new_path='path/to/files/new'
    switch_dict={'old':'new'}
    
    def __init__(self,dictionary):
        for k,v in dictionary.items():
            setattr(self,k,v)
        self.BRANCH=self.BRANCH.split('#')[1]
        self.MEETING=self.MEETING.strip()
        self.mtg_date=get_dates(self.MEETING)
        self.version=get_version(self.Name)
        self.access=get_access(self.Name)
        
    def get_recipients(self,cpm_data):
        recipients=[placeholders here]
        for d in cpm_data:
            if d['Branch']==self.cpm_branch:
                recipients.append(d['CPM'])
        self.recipients=recipients
    
    def get_uploading_name(self):
        if self.ProtocolName in self.Name:
            self.uploading_name=self.Name
        else:
            self.uploading_name=self.ProtocolName+'_'+self.Name
    
    def associate_protocol(self):
        for k,v in document.protnum_associations.items():
            if self.ProtocolName in v:
                self.ProtocolGroup = k
            else:
                self.ProtocolGroup = self.ProtocolName
    
    def get_mtg_data(self,meeting_dict):
        if self.mtg_date is np.nan:
            self.mtg_name = 'select'
            self.mtg_id = 0
        else:
            self.mtg_check1 = self.BRANCH+'-'+self.ProtocolName+'-'+self.cmt_type+' '+self.mtg_date
            for d in meeting_dict:
                if date_matcher(self.mtg_check1)==date_matcher(d['MeetingName']):
                    self.mtg_name=d['MeetingName']
                    self.mtg_id=d['MeetingID']
                    break
                else:
                    if self.ProtocolName in document.protnum_associations:
                        for i in document.protnum_associations[self.ProtocolName]:
                            self.mtg_check = self.BRANCH+'-'+i+'-'+self.cmt_type+' '+self.mtg_date
                            if date_matcher(self.mtg_check) == date_matcher(d['MeetingName']):
                                self.mtg_name = d['MeetingName']
                                self.mtg_id  = d['MeetingID']
                                break
                    else:
                        self.mtg_name = 'select'
                        self.mtg_id = 0
    
    def get_mtg_button_type(self):
        if "Minute" in self.DOC_TYPE or "Recomm" in self.DOC_TYPE:
            self.mtg_button_type='Post-Meeting'
        else:
            self.mtg_button_type='Meeting'
            
    #functions to remove special characters and rename paths if necessary
    def clean_uploading_name(self):
        exclude="!@#$%^&*()[]{};:,/<>?\|`~'=+"
        if any(i in self.uploading_name for i in exclude):
            self.uploading_name=self.uploading_name.translate({ord(c): "" for c in exclude})
        else:
            pass
 
    def get_location(self):
        self.location=document.old_path+'/'+self.PROTOCOL+'/'+self.MEETING+'/'+self.Name
    
    def get_clean_location(self):
        self.clean_location=document.old_path+'/'+self.PROTOCOL+'/'+self.MEETING+'/'+self.uploading_name
    
    def get_destination(self):
        yr_part=re.search('(\d{2})\-\d{3,4}',self.PROTOCOL)
        if yr_part:
            self.destination=document.new_path+'/'+yr_part[1]+self.PROTOCOL
        else:
            if self.PROTOCOL in document.switch_dict.keys():
                self.destination=document.new_path+'/'+yr_part[1]+self.PROTOCOL
    
    # functions to bundle protocol associated/mtg associated docs
    def get_posting_association(self):
        if 'email' in self.Name or " em " in self.Name:
            self.association='none'
        else:
            if int(self.mtg_id)==0:
                self.association='protocol'
            if int(self.mtg_id)>0:
                self.association='meeting'

# metadata of documents to upload
target_file='documents_to_upload.xlsx'

# load and clean needed data
datapath = 'path/to/datadump/'
upload_posting_folder_test = 'path/to/test/'
upload_posting_folder = 'path/to/posting_uploading/'
df = pd.read_excel(upload_posting_folder+target_file)
df = df[~df['BRANCH'].str.contains('unknown')]
df = df[['Name','DOC_GROUP','DOC_TYPE','PROTOCOL','BRANCH','MEETING']]
dl_to_protnames = {'name1':'name2'}
df['ProtocolName'] = df['PROTOCOL'].replace(dl_to_protnames,regex = True)
df_cmt = pd.read_excel(upload_posting_folder_test+'prot_cmt_type.xlsx')
df_cmt.columns = ['ProtocolName','cmt_type']
df_cpm = pd.read_excel(upload_posting_folder_test+'cpm_branch.xlsx')
cpm_dict = df_cpm.to_dict('records')
df_branch_adjusted = pd.read_excel(datapth + "prot_branch_adjusted.xlsx")
df_branch_adjusted.columns = ['ProtocolName','cpm_branch']
branch_adjusted_dict = df_branch_adjusted.to_dict('records')
df_prot_title = pd.read_excel(datapath+'prot_fulltitle.xlsx')
title_dict = df_prot_title.to_dict('records')
df1 = df.merge(df_cmt,on = 'ProtocolName',how = 'left')
df1['mtg_date'] = df1['MEETING'].apply(lambda x: get_dates(x))
df1=df1.merge(df_branch_adjusted,on='ProtocolName',how='left')
docs_final_dict = df1.to_dict('records')
uploaded_file_protocols = df1['ProtocolName'].tolist()

# JSON of meeting data to associate with documents to-be-migrated
with open(datapath+"jsons/meeting_json_dump.txt") as json_data:
    mtg_dict=json.load(json_data)
output_dict1=[x for x in mtg_dict["Data"]]
mtg_dict_narrowed=deduplicate_filtered_keys(
    filter_keys(
        output_dict1,['MeetingID','MeetingName']
    )
)

#clean and bundle document data into dictionaries and append to a list of dictionaries
to_upload=[]
for doc in docs_final_dict:
    try:
        o=document(doc)
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

# upload documents to new web-based document repository and append dicts of successfully uploaded docs to uploaded_docs list
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
