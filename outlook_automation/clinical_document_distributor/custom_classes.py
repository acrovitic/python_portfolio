class protocol_doc(object):
    def __init__(self,url):
        self.url = url
        self.protocol = re.search('(\d{2}\-\d{4})',url)[0]
        self.path = download_protocol(self.url) #download doc, returns m drive location downloaded to
        self.team = get_team_personnel(self.protocol)
     
    def get_new_doc_attributes(self): # reads downloaded document and retrieves key data points
        fpart = re.search('Protocol\/(.*?)\.',self.path)[1]
        text = docx2txt.process(self.path)
        text1 = [i for i in text.splitlines() if len(i)>1]
        self.date = get_dates(find_item(text1, lambda i: get_dates(i)))
        self.version = get_version(find_item(text1, lambda i: "ver" in i.lower()))
        self.title = get_title(text1, lambda i: "title:" in i.lower())
        self.new_file_name = fpart+'_'+self.version+"_"+self.date
        self.new_file_path = self.path.replace(fpart,self.new_file_name)

class doc_package(object):
    def __init__(self,dictionary):
        for k,v in dictionary.items():
            setattr(self,k,v)
        self.paired_paths = list(zip(self.path_,self.new_file_path_))
    
    def rename_downloaded_files(self):
        for i in self.paired_paths:
            if os.path.isfile(i[0]):
                os.rename(i[0],i[1])
            else:
                print("Not found.")
    
    def set_email_msg_part(self):
        if "redline" and "changes" in ' '.join(self.new_file_path_).lower():
            self.msg_part = ', along with the Redline version and Summary of Changes, for Charter development.'
        elif any("redline" in i.lower() for i in self.new_file_path_):
            self.msg_part = ', along with the Redline version, for Charter development.'
        elif any("changes" in i.lower() for i in self.new_file_path_):
            self.msg_part = ', along with the Summary of Changes, for Charter development.'
        else:
            self.msg_part = ' for Charter development.'
    
    def send_icon_new_docs(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = self.icon[1]
        mail.CC = 'groupinbox@example.com;'
        mail.Subject = email_subject.format(protocol=self.protocol, version=self.version)
        for item in self.new_file_path_:
            mail.Attachments.Add(item)
        mail.HTMLBody=email_template.format(protocol=self.protocol,title=self.title,
                                            icon=self.icon[0],version=self.version,
                                            date=self.date,msg_part=self.msg_part)
        mail.Display(False)
        
    def move_docs_on_mdrive(self):
        template_1study_path = "path/to/studies/20{y}/{p}/Essential Documents/{f}"
        exclude = ['redline','change']
        for item in self.new_file_path_:
            if not any(x in item.lower() for x in exclude):
                yr_part = re.search('(\d{2})\-\d{4}',self.protocol)[1]
                posted_path = template_1study_path.format(y=yr_part,p=self.protocol,f=item.rsplit("/",1)[1])
                self.posted_path = posted_path
                copyfile(item,self.posted_path)
                print("Moved new Protocol to {np}.\n------".format(np=self.posted_path))
                pfolder = self.posted_path.rsplit("/",1)[0]+"/"
                protnum = re.search('(\d{2}\-\d{4})',pfolder)[0]
                list_of_protfiles = glob.glob(pfolder+"*Protocol*")
                if len(list_of_protfiles)>1:
                    old_protocol_loc = min(list_of_protfiles, key=os.path.getctime)
                    old_protocol_dest = old_protocol_loc.replace("Posted","Removed")
                    try:
                        os.rename(old_protocol_loc,old_protocol_dest)
                    except:
                        print("Could not move {f}.\n-----".format(f=self.posted_path.rsplit("/",1)[1]))
                else:
                    print("Only latest version of protocol in {pn}.\n-----".format(pn=self.protocol))
                    continue
                
    def get_posting_metadata(self,bdict):
        self.access = "Open"
        self.doc_type = "Protocol"
        self.status = "Final"
        self.old_dl_mtg = '1 - Study Documents'
        self.old_dl_doctype = 'Protocol'
        for d in bdict:
            if self.protocol==d['Protocol Number']:
                    self.Branch=d['Branch']
