import win32com
from win32com import client
import datetime as dt
import sys, os, re
import pandas as pd
import time
from dateutil.parser import parse
from collections import Counter

# make one dictionary for drive filing path and one dictionary for outlook filing path.
m_dict = dict() # path for email on m drive
m_attachment = dict() # path for attachments on m drive
outlook_dict = dict() # path for outlook folder
date = dt.datetime.now().strftime("%m/%d/%Y") # today
for email in list(main_folder.Folders["Inbox"].Items.restrict("[SentOn] < '{} 07:00 AM' And [Categories] = 'File'".format(date))):
    match = re.search("(\d{2})\-\d{4}",email.Subject)
    if match:
        try:
            prot_num = match.group(0) # one_tuple previously used to get protocol number
            year_part = match.group(1)
            file_path = "path/to/protocol_folders/20" + year_part + "/" + prot_num + "/Emails/"
            attachment_path = "path/to/protocol_folders/20" + year_part + "/" + prot_num + "/Emails/Attachments/"
            strname = str(email.Subject).replace(":","-")
            outlook_path = main_folder.Folders["Inbox"].Folders[prot_num].Folders["Completed"]
            m_dict[str(email.Subject)] = file_path+strname+".msg"
            m_attachment[str(email.Subject)] = attachment_path
            outlook_dict[str(email.Subject)] = outlook_path
        except:
            print(strname)
            break

# save off attachments to their appropriate protocol folder before filing emails
m = []
a = []
o = []
date = dt.datetime.now().strftime("%m/%d/%Y") # today
# saves attachments
for email in list(main_folder.Folders["Inbox"].Items.restrict("[SentOn] < '{} 07:00 AM' And [Categories] = 'File'".format(date))):
    for k,v in m_attachment.items():
        if str(email.Subject) in k:
            attpath = v
            if email.Attachments.Count > 0:
                for att in email.Attachments:
                    if not os.path.exists(attpath):
                        os.makedirs(attpath)
                    att.SaveAsFile(attpath+att.FileName)
                    print("new path made for {}".format(str(att.FileName)))

# file off emails
date = dt.datetime.now().strftime("%m/%d/%Y") # today
# saves attachments
for email in list(main_folder.Folders["Inbox"].Items.restrict("[SentOn] < '{} 07:00 AM' And [Categories] = 'File'".format(date))):
    for k,v in m_dict.items():
        if str(email.Subject) in k:
            try:
                email.SaveAs(v)
                print("{} saved to M.".format(str(email.Subject)))
            except:
                print("m filing error with {}".format(str(email.Subject)))
                break
    for k,v in outlook_dict.items():
        if str(email.Subject) in k:
            try:
                email.Move(v)
                print("{} moved in Outlook.".format(str(email.Subject)))
            except:
                print("outlook filing error with {}.".format(str(email.Subject)))
                break
