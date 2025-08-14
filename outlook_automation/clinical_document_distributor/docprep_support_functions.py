# functions for posting to dl and m
# required custom functions
def get_doc_attributes(string):
    string=str(string)
    study_docs=['Charter']
    mtg_docs=['Recom','Minute']
    mtg_types=['Org','Data Review','DRM','Hoc','E-Rev','Electronic Rev']
    path_template="path/to/studies/"
    prot_num=re.search('(\d{2})\-\d{4}',string)
    date=re.search('(\d{2}\w{3}\d{2})',string)
    if any(i in string for i in study_docs):
        return [path_template+"/Essential Documents/",string]
    elif any(i in string for i in mtg_docs):
        for i in mtg_types:
            if i in string:
                return [path_template+"/Meetings",date[0],string]
    else:
        return "Unknown"


def get_mtg_path(attr_list):
    if attr_list:
        meeting_folders=os.listdir(attr_list[0])
        for folder in meeting_folders:
            date=parse(folder,fuzzy=True)
            if date==parse(attr_list[1],fuzzy=True):
                return ''.join(attr_list[0])+"/"+folder+"/"
            else:
                return "no matching meeting folder found"
    else:
        return "Error: Attribute is 'None' type."


def get_doc_type(file_path):
    type_dict={'Recommendations':['Recom'],#add more later
           'Meeting Minutes - Open':['Minute','Open'],
           'Meeting Minutes - Closed':['Minute','Closed'],
           'Charter':['Charter']}
    file_name=file_path.rsplit('/',1)[1]
    for k,v in type_dict.items():
        if all(i in file_name for i in v):
            return k


def old_dl_mtg_format(string):
    if 'Charter' not in string:
        meeting=string.rsplit('/',2)[1]
        date=parse(meeting.split(' ',1)[1].replace("DRM ","")).strftime('%d %B %Y')
        type_dict={'Data Review Meeting':['DRM','Data Review'],
                   'Organizational Meeting':['Org','Organizational'],
                   'Ad-Hoc Meeting':['Ad hoc','Ad-hoc','Ad-Hoc','Ad Hoc']}
        for k,v in type_dict.items():
            if any(i in meeting for i in v):
                return k+' - '+date
    else:
        return "1 - Study Documents"


def detect_date(file_path):
    mtg_part=file_path.rsplit("/",2)[1]
    for item in mtg_part.split():
        try:
            parsed_item=parse(item)
            return parsed_item.strftime('%m/%d/%Y')
        except ValueError:
            continue


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
