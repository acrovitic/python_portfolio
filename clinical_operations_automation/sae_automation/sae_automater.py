# modules
import win32com
import datetime as dt
import os
import re
import datetime as dt
from sae_automater_templates import sae_initial_email, sae_followup_email, sae_init_fu_email

# select shared inbox folder 
outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
main_folder=outlook.Folders[2] 
inbox_emails=main_folder.Folders["Inbox"].Items

day=dt.datetime.now()+dt.timedelta(days=7)
day=day.strftime('%d-%b-%y')

# functions
def get_data_point(email_body,datapoint):
    datapoint=str(datapoint)
    data_point_line=''.join([i for i in email_body if datapoint in i])
    match=re.search(re.escape(datapoint)+': (.*)</font>',data_point_line)
    return match[1]
    
def remove_dupes(l, k):
    seen = {} 
    for d in l:
        v=d[k]
        if v not in seen:
            seen[v] = d
    return seen

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

#sae email handling class
class sae_email(object):
    fu_subj="Safety Oversight, Protocol {prot}, SAE Follow-up #{funum} Notification, Subject {sid} ({cond})"
    init_subj="Safety Oversight, Protocol {prot}, SAE Initial Notification, Subject {sid} ({cond})"
    ordinal=lambda x:["initial","first","second","third","fourth","fifth","sixth","seventh","eighth","ninth",
                      "tenth","eleventh","twelfth"][x-1]
    
    def __init__(self,dictionary):
        for k,v in dictionary.items():
            setattr(self,k,v)
        self.nbodies=[] #sae notification email bodies
        self.rbodies=[] #sae report email bodies
        self.nformatted=[] #format sae notices
        self.rformatted=[] #format report
    
    def get_type(self):
        if "Initial" in self.notices[0] and "Initial" in self.reports[0]:
            self.type="Initial"
        else:
            p=re.search('(\d{2}\-\d{4})\-',self.sae_id)[1]
            r=[self.prot,'Initial ISM SAE Report',self.subjectid,self.condition]
            s=[self.prot,self.subjectid,self.condition]
            lim=int(len(list(main_folder.Folders["Inbox"].Folders[p].Folders['Completed'].Items))/2)
            for email in list(main_folder.Folders["Inbox"].Folders[p].Folders['Completed'].Items)[:lim]:
                if all([i in str(email.Subject) for i in s]):
                    self.type="Follow-up Reminder"
                else:
                    self.type="Follow-up"
    
    def extract_attributes(self):
        self.prot=re.search('(\d{2}\-\d{4})',self.sae_id)[0]
        self.subjectid=self.notices[0].split(", ",2)[1].replace("Subject ID ","")
        self.condition=self.reports[-1].rsplit(", ",2)[1]
        self.day=(dt.datetime.now()+dt.timedelta(days=7)).strftime('%d-%b-%y')

    def get_followup_nums(self):
        if "Follow-up" not in self.type:
            pass
        else:
            self.funums=[]
            for ele in self.reports:
                match=re.search('Follow-up #(\d)',ele)
                if match:
                    self.funums.append(int(match[1]))
                self.followup=max(self.funums)
    
    def get_followup_history(self):
        if "Initial" not in self.type:
            self.funums.sort()
            x=[i for i in range(1,min(self.funums))]
            if len(x)==0:
                self.fuhist="initial"
            if len(x)>0:
                x.insert(len(x)-1,"and")
                self.fuhist=', '.join(x[:len(x)-2])+' '+' '.join(x[len(x)-2:])

    def get_single_plural(self):
        if len(self.reports)>1:
            self.spx=["notifications","Narrative Summaries","have","are"]
        if len(self.reports)==1:
            self.spx=["notification","Narrative Summary","has","is"]
        self.sp1=self.spx[0]
        self.sp2=self.spx[1]
        self.sp3=self.spx[2]
        self.sp4=self.spx[3]
    
    def get_ordinals(self):
        if "Follow-up" in self.type:
            self.funums.sort()
            self.fnords=[sae_email.ordinal(i+1) for i in self.funums]
            self.fnords.insert(len(self.fnords)-1,"and")
            self.ords=', '.join(self.fnords[:len(self.fnords)-2])+' '+' '.join(self.fnords[len(self.fnords)-2:])
   
    def get_body_content(self):
        for email in list(main_folder.Folders["Inbox"].Items.Restrict("[Categories]='In Process'")):
            if str(email.Subject) in self.notices:
                self.nbodies.append(email.HTMLBody)
            if str(email.Subject) in self.reports:
                self.rbodies.append(email.HTMLBody)
    
    def format_notices(self):
        i=1
        if "Follow-up" in self.type:
            self.nbodies.reverse()
            for notice in self.nbodies:
                split_notice=notice.split('<p>')
                title=''.join([x for x in split_notice if "Protocol Title" in x])
                match=re.search("Protocol Title: (.*)</font>",title)
                self.title=match[1]
                self.header="<u><strong>Follow-up #{followup_number} SAE Notification</strong></u><br>".format(followup_number=i)
                formatted_notice='<p>'.join(split_notice[3:12])
                self.nformatted.append(self.header+formatted_notice)
                i+=1
        if self.type=="Initial":
            for notice in self.nbodies:
                split_notice=notice.split('<p>')
                title=''.join([x for x in split_notice if "Protocol Title" in x])
                match=re.search("Protocol Title: (.*)</font>",title)
                self.title=match[1]
                self.header="<u><strong>Initial SAE Notification</strong></u><br>"
                formatted_notice='<p>'.join(split_notice[3:12])
                self.nformatted.append(self.header+formatted_notice)
    
    def format_reports(self):
        i=1
        exclude=[ 'MM: To complete the Approval process, please log into:',
                 'Links to Investigator Brochure:','Links to Protocol:']
        if "Follow-up" in self.type:
            self.rbodies.reverse()
            for report in self.rbodies:
                split_report=[x for x in report.split('<p>')[5:] if 'medicalresearch.com' not in x]
                split_report1=[x for x in split_report if not any(j in x for j in exclude)]
                del split_report1[-1:]
                self.header="<u><strong>Follow-up #{followup_number} SAE Report</strong></u><br>".format(followup_number=i)
                report_msg='<p>'.join(split_report1)
                self.rformatted.append(self.header+report_msg)
                i+=1
        if "Initial" in self.type:
            for report in self.rbodies:
                split_report=[x for x in report.split('<p>')[5:] if 'medicalresearch.com' not in x]
                split_report1=[x for x in split_report if not any(j in x for j in exclude)]
                self.header="<u><strong>Initial SAE Report</strong></u><br>"
                report_msg='<p>'.join(split_report1)
                self.rformatted.append(self.header+report_msg)

    def get_subject(self):
        if "Follow-up" in self.type:
            if len(self.funums)==1:
                self.subject=sae_email.fu_subj.format(prot=self.prot,
                                                   funum=str(self.funums[0]),
                                                   sid=self.subjectid,
                                                   cond=self.condition)
            if len(self.funums)>1:
                self.funums.sort()
                self.funums.insert(len(self.funums)-1,"and")
                self.funums=[str(i) for i in self.funums]
                self.funum_subj=', '.join(self.funums[:len(self.funums)-2])+' '+' '.join(self.funums[len(self.funums)-2:])
                self.subject=sae_email.fu_subj.format(prot=self.prot,
                                                   funum=self.funum_subj,
                                                   sid=self.subjectid,
                                                   cond=self.condition)
                self.subject=self.subject.replace("Notification","Notifications")
                
        if self.type=="Initial":
            self.subject=sae_email.init_subj.format(prot=self.prot,
                                                 sid=self.subjectid,
                                                 cond=self.condition)
    
    def create_email(self):
        if self.type=="Follow-up":
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'johndoe@example.com'
            mail.Subject = self.subject
            mail.HTMLBody = sae_followup_email.format(prot=self.prot,
                                                      title=self.title,
                                                      sp1=self.sp1,
                                                      sp2=self.sp2,
                                                      sp3=self.sp3,
                                                      sp4=self.sp4,
                                                      day=self.day,
                                                      subject_id=self.subjectid,
                                                      condition=self.condition,
                                                      followup_ord=self.ords,
                                                      notification="<br><br>".join(self.nformatted),
                                                      report="<br><br>".join(self.rformatted))
            mail.Display(False)
        if self.type=="Follow-up Reminder":
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'johndoe@example.com'
            mail.Subject = self.subject
            mail.HTMLBody = sae_init_fu_email.format(prot=self.prot,
                                                      title=self.title,
                                                      fuhist=self.fuhist,
                                                      sp1=self.sp1,
                                                      sp2=self.sp2,
                                                      sp3=self.sp3,
                                                      sp4=self.sp4,
                                                      day=self.day,
                                                      subjectid=self.subjectid,
                                                      condition=self.condition,
                                                      followup_ord=self.ords,
                                                      notification="<br><br>".join(self.nformatted),
                                                      report="<br><br>".join(self.rformatted))
            mail.Display(False)
        if self.type=="Initial":
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'johndoe@example.com'
            mail.Subject = self.subject
            mail.HTMLBody = sae_initial_email.format(prot=self.prot,
                                                      title=self.title,
                                                      day=self.day,
                                                      subjectid=self.subjectid,
                                                      condition=self.condition,
                                                      notification="<br><br>".join(self.nformatted),
                                                      report="<br><br>".join(self.rformatted))
            mail.Display(False)

# get unique sae id's
sae_id_list=[]

for email in list(main_folder.Folders["Inbox"].Items.Restrict("[Categories]='In Process'")):
    match=re.search("(\d{2}\-\d{4}\-\d{5})",email.Subject)
    if match[0] not in sae_id_list:
        sae_id_list.append(match[0])

# get sae id specific notification/report email subjects
l=[]
for sae_id in sae_id_list:
    d={}
    d['sae_id']=sae_id
    notices=[]
    reports=[]
    for email in list(main_folder.Folders["Inbox"].Items.Restrict("[Categories]='In Process'")):
        if sae_id in str(email.Subject) and 'Notification' in str(email.Subject):
            notices.append(str(email.Subject))
        if sae_id in str(email.Subject) and 'Report' in str(email.Subject):
            reports.append(str(email.Subject))
    d['notices']=notices
    d['reports']=reports
    l.append(d)

final_sae_email_dict=deduplicate_filtered_keys(
    filter_keys(
        l,['sae_id','notices','reports']
    )
)

for sae_d in final_sae_email_dict:
    if len(sae_d['reports']) > 0:
        s=sae_email(sae_d)
        s.extract_attributes()
        s.get_type()
        s.get_body_content()
        print("Body content extracted.")
        s.format_notices()
        s.format_reports()
        print("Message content formatted.")
        s.get_followup_nums()
        s.get_followup_history()
        s.get_single_plural()
        s.get_ordinals()
        s.get_subject()
        print("Creating email.")
        s.create_email()
