
import os
import re
import glob
import getpass
import difflib
import numpy as np
import pandas as pd
from email_templated import prior, noprior
from docx import Document
from datetime import date
from mailmerge import MailMerge
import win32com.client as win32
from datetime import datetime, timedelta

user = getpass.getuser().upper()

welcome_message = """
UTILITY: Invitation Package Assembler.
USER: {}
INSTRUCTIONS: 
To generate an invitation package, provide requested information in the order specified below. 
protocol#;first name;last name;role;prior/noprior-mid/start;month, year of anticipated mtg or x if no date yet available
Once provided, press ENTER.
Pressing ENTER after providing required information for a package will generate a new prompt to create another package. You can create as many as required.
When finished, press ENTER on an empty prompt to initiate invitation generation.
"""

class invitee(object):
    mpath_template = "pth/to/study/folder/20{y}/{p}/"
    def __init__(self,info_list):
        self.protocol = info_list[0]
        self.first_name = info_list[1]
        self.last_name = info_list[2]
        self.name_key0 = self.first_name + ' ' + self.last_name
        self.role = info_list[3]
        self.template_flag = info_list[4]
        self.anticipated_mtg = info_list[5]
        self.etype = self.template_flag.split("-")[0]
        self.yr_part = re.search('(\d{2})\-\d{4}',self.protocol)[1]
        self.coi_mpath1 = glob.glob(invitee.mpath_template.format(y=self.yr_part,p=self.protocol)+"* Info")[0]
        self.coi_mpath2 = glob.glob(invitee.mpath_template.format(y=self.yr_part,p=self.protocol)+"* Info"+"/COIs")[0]
        self.summary_mpath = invitee.mpath_template.format(y=self.yr_part,p=self.protocol)+"Essential Documents/Protocol/"
        self.Role = {'Chair': 'the Chair',
                            'Member':'a Member',
                            'Biostatistician':'the Biostatistician'}[self.role]
        self.name_key = (self.name_key0[:3] + '.' + self.name_key0[self.name_key0.find(' '):]).lower()
        self.due_date = due_date = datetime.now() + timedelta(days=7)
        self.due_date_formatted = self.due_date.strftime('%d-%b-%y')
    
    def get_study_attrs(self,prot_dict): # protocol_dict
        for k,v in [d for d in prot_dict if d['Protocol_Number'] == self.protocol][0].items():
            setattr(self,k,v)
        self.Committee = {'SMC':'Safety Monitoring Committee (SMC)',
                              'DSMB':'Data and Safety Monitoring Board (DSMB)'}[self.Committee_Type]
        self.sentence = {'SMC':'an SMC',
                           'DSMB':'a DSMB'}[self.Committee_Type]
    
    def get_anticipated_mtg(self):
        if self.anticipated_mtg.lower() == "x":
            if self.template_flag.split("-")[1].lower() == "start":
                self.mtgsnt = "An Organizational Meeting teleconference is anticipated to be held in the near future."
            if self.template_flag.split("-")[1].lower() == "mid":
                self.mtgsnt = "A Data Review Meeting teleconference is anticipated to be held in the near future."
        if self.anticipated_mtg.lower() != "x":
            if self.template_flag.split("-")[1].lower() == "start":
                self.mtgsnt = "It is anticipated that there will be {a} Organizational Meeting teleconference in {b}.".format(a=self.sentence,
                                                                                                                              b=self.anticipated_mtg)
            if self.template_flag.split("-")[1].lower() == "mid":
                self.mtgsnt = "It is anticipated that there will be {a} Data Review Meeting teleconference in {b}.".format(a=self.sentence,
                                                                                                                              b=self.anticipated_mtg)
    
    def get_contact_info(self,welc_dict):
        all_lengths = [] # hold length of all values for all dicts with same name_key (reason: duplicates exist in cms)
        for d in welc_dict:
            if self.name_key == d['name_key']:
                d['total_length'] = 0
                for k,v in d.items():
                    d['total_length'] += len(str(v))
                all_lengths.append(d['total_length'])
                if d['total_length'] == max(all_lengths):
                    for k,v in d.items():
                        setattr(self,k,v)
    def set_template_names(self):
        self.letter = self.Protocol_Number+' '+self.Committee_Type+' Welcome Letter - '+self.Last_Name+'.docx'
        self.cif = self.Protocol_Number+' Contact Information Form - '+self.Last_Name+'.docx'
        self.subj = 'Safety Oversight, Protocol '+self.Protocol_Number+', '+self.Committee_Type+' Membership Invitation - '+self.Last_Name

    def get_coi_forms(self):
        if any(self.last_name.lower() and ".pdf" in i.lower() for i in os.listdir(self.coi_mpath1)):
            for f in os.listdir(self.coi_mpath1):
                if self.last_name.lower() in f.lower() and f.endswith('.pdf'):
                    self.coi = os.path.join(self.coi_mpath1,f)
        elif any(self.last_name.lower() and ".pdf" in i.lower() for i in os.listdir(self.coi_mpath2)):
            for f in os.listdir(self.coi_mpath2):
                if self.last_name.lower() in f.lower() and f.endswith('.pdf'):
                    self.coi = os.path.join(self.coi_mpath2,f)
        else:
            self.coi = 'not available'
    
    def get_protocol_summary(self):
        list_of_files = glob.glob(self.summary_mpath+"/*Summary*")
        if len(list_of_files) < 1:
            self.protocol_summary = 'not available'
        else:
            self.protocol_summary = max(list_of_files, key=os.path.getctime)

if __name__ == '__main__':
    path1 = os.getcwd() + "\\data\\"
    path2 = os.getcwd() + "\\templates\\"
    list_of_welcomepackage_files = glob.glob(path1 + 'welcomepackage*')
    latest_welcome_package = max(list_of_welcomepackage_files, key=os.path.getctime)
    xls1 = pd.ExcelFile(latest_welcome_package)
    welcome_df = xls1.parse(xls1.sheet_names[0])
    welcome_df.fillna('',inplace=True)
    welcome_dict = welcome_df.to_dict('records')
    for d in welcome_dict:
        d['name_key'] = (d['First_Name'][:3] + '. ' + d['Last_Name']).lower()
    list_of_prot_files = glob.glob(path1 + 'CMSReport-prots*')
    latest_prot_file = max(list_of_prot_files, key=os.path.getctime)
    xls2 = pd.ExcelFile(latest_prot_file)
    protocol_dict = xls2.parse().to_dict('records')
    print(welcome_message.format(user))

    i=0 
    name_role = []
    while 1:
        i+=1
        name=input("Candidate %d: \n" %i)
        if name=='':
            break
        inv = invitee([ele.strip() for ele in name.split(';')])
        inv.get_study_attrs(protocol_dict)
        inv.get_contact_info(welcome_dict)
        inv.set_template_names()
        inv.get_anticipated_mtg()
        inv.get_coi_forms()
        inv.get_protocol_summary()
        name_role.append(vars(inv))
    cif = path2+'Template Contact Information Form.docx'
    outpath = 'path/to/script/output/'
    email_types = {'noprior': noprior} # add prior email template and mid/start templates
    for d in name_role:
    	welcome_email = email_types[d['etype']]
    	welcome_letter = path2+'Template Welcome Letter-{}.docx'.format(d['template_flag'])
    	doc_letter = MailMerge(welcome_letter)
    	doc_cif = MailMerge(cif)
    	name1 = outpath+d['letter']
    	name2 = outpath+d['cif']
    	print('Package complete for',d['Last_Name'])
    	doc_letter.merge_pages([d])
    	doc_cif.merge_pages([d])
    	doc_letter.write(name1)
    	doc_cif.write(name2)
    	doc_letter.close()
    	doc_cif.close()
    	outlook = win32.Dispatch('outlook.application')
    	recipient = d['Last_Name']
    	prot = d['Protocol_Number']
    	titl = d['Protocol_Full_Title']
    	cmt = d['Committee']
    	cmt_abrv = d['Committee_Type']
    	due_date = datetime.now() + timedelta(days=7)
    	due_date_formatted = due_date.strftime('%d-%b-%y')
    	document1 = Document(outpath+d['letter'])
    	core_properties1 = document1.core_properties
    	core_properties1.author = d['Protocol_Number']
    	document1.save(outpath+d['letter'])
    	document2 = Document(outpath+d['cif'])
    	core_properties2 = document2.core_properties
    	core_properties2.author = d['Protocol_Number']
    	document2.save(outpath+d['cif'])
    	mail = outlook.CreateItem(0)
    	mail.To = d['Email_Address1']
    	mail.CC = 'SOCS@dmidcroms.com; '+d['Assistant_Email']
    	mail.Subject = d['subj']
    	mail.Attachments.Add(outpath+d['letter'])
    	mail.Attachments.Add(outpath+d['cif'])
    	if d['coi'] != 'not available':
    		mail.Attachments.Add(d['coi'])
    	else:
    		continue
    	if d['protocol_summary'] != 'not available':
    		mail.Attachments.Add(d['protocol_summary'])
    	else:
    		continue
    	mail.HTMLBody = welcome_email.format(**d)
    	mail.Display(False)
