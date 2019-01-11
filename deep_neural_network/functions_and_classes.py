  # functions to find top keywords
def prepare_body(string):
    for r in (("\r\n",""),("\t",""),("\n","")):
        string=string.replace(*r)
    return string

def get_sender_text(string):
    if "/o=" in string:
        return string.rsplit("-",1)[1]
    if "/O=" in string:
        return string.rsplit("=",1)[1]
    else:
        return string
    
def parse_email(email_text, remove_quoted_statements=True):
    email_text = email_text.strip()
    email_text = strip_automated_notation(email_text)    
    if remove_quoted_statements:
        pattern = """(?P<quoted_statement>".*?")"""
        matches = re.findall(pattern, email_text, re.IGNORECASE + re.DOTALL)
        for m in matches:
            email_text = email_text.replace(m, '"[quote]"')
    result = { \
              "salutation":get_salutation(email_text), \
              "body":get_body(email_text), \
              "signature":get_signature(email_text), \
              "reply_text":get_reply_text(email_text) \
              }
    return result

def get_body(email_text, check_salutation=True, check_signature=True, check_reply_text=True):
    
    if check_salutation:
        sal = get_salutation(email_text)
        if sal: email_text = email_text[len(sal):]
    
    if check_signature:
        sig = get_signature(email_text)
        if sig: email_text = email_text[:email_text.find(sig)]
    
    if check_reply_text:
        reply_text = get_reply_text(email_text)
        if reply_text: email_text = email_text[:email_text.find(reply_text)]
            
    return email_text
    
def clean_text_content(content):
    c=re.sub('<[^>]+>', '', content)
    c=re.sub('(Contractor).*?\r\n\r\n \r\n\r\n \r\n\r\n \r\n\r\n','',c, flags=re.DOTALL)
    c=re.sub('(Contractor).*?Ireland','',c, flags=re.DOTALL)
    c=re.sub('(Contractor).*?please notify us immediately at GROUP@example.com','',c, flags=re.DOTALL)
    c=re.sub('From: .*?Subject:','',c, flags=re.DOTALL)
    return c

def clean_up_sentence(sentence):
    # tokenize the pattern - split words into array
    sentence_words = nltk.word_tokenize(sentence)
    # stem each word - create short form for word
    sentence_words = [stemmer.stem(word.lower()) for word in sentence_words]
    return sentence_words

# return bag of words array: 0 or 1 for each word in the bag that exists in the sentence
def bow(sentence, words, show_details=True):
    # tokenize the pattern
    sentence_words = clean_up_sentence(sentence)
    # bag of words - matrix of N words, vocabulary matrix
    bag = [0]*len(words)  
    for s in sentence_words:
        for i,w in enumerate(words):
            if w == s: 
                # assign 1 if current word is in the vocabulary position
                bag[i] = 1
                if show_details:
                    print ("found in bag: %s" % w)

    return(np.array(bag))

  # outlook-specific functions, classes, etc
# functions, text templates, and category tokens
def build_outlook_path(predicted_path):
    f='.Folders[{}]'*len(predicted_path.split(';'))
    l=[]
    for a,b in list(zip(f.split('.')[1:],predicted_path.split(';'))):
        b="'"+b+"'"
        c=a.format(b)
        l.append(c)
    if "SARFs" not in 'main_folder.'+'.'.join(l):
        return 'main_folder.'+'.'.join(l)
    else:
        return 'main_folder.'+'.'.join(l)+'.Folders[YYYY]'
def format_outlook_path(path_string,rcvd):
    if "yy" in path_string.lower():
        date_parts = [rcvd.strftime("%m"),rcvd.strftime("%B"),rcvd.strftime("%Y")]
        if "TEAM" in path_string:
            for r in (("dd","*"+date_parts[0]),('Mmmm',date_parts[1]),("yyyy",date_parts[2])):
                path_string = path_string.replace(*r)
            return path_string
        else:
            return path_string.replace("YYYY","*"+date_parts[2])
    else:
        return path_string
    
# template text
c_ai_text = "C_AI[READING]:{e}\nC_AI[CATEGORIZED]:{p}\n------"
f_ai_text = "F_AI[READING]:{e}\nF_AI[LOCATION]:{f}\n------"
f_ai_text2 = "F_AI[READING]:{e}\nF_AI[LOCATION]:{f}\nF_AI[M]:{m}\nF_AI[M:ATT]:{a}\n------"

class email_obj(object):
    def __init__(self,email):
        self.body = email.Body
        self.subject = email.Subject
        try:
            self.sender = email.Sender.Address
        except:
            self.sender = "None"
        try:
            self.to = email.To
        except:
            self.to = "None"
        self.received = email.ReceivedTime
        self.attcount = email.Attachments.Count
        if not email.Categories:
            self.category = 'none'
        else:
            self.category = email.Categories
            if self.category == "File":
                self.prot = re.search('(\d{2}\-\d{4})',self.subject)[0]
                self.yr_part = self.prot.split("-")[0]
                self.outlook = "Inbox;{p};Completed".format(p=self.prot)
                self.mdrive = "custom/path/to/study/{}/Emails/".format(y=self.yr_part, p=self.prot)
                self.attpath = "custom/path/to/study/{}/Emails/Attachments/".format(y=self.yr_part, p=self.prot)
        
    def clean_attributes(self):
        x = parse_email(clean_text_content(prepare_body(str(self.body))))['body']
        x = re.sub('\w+:([^.]*)*','',x)
        x = re.sub('[^a-zA-Z]+', ' ', x)
        self.body = x
        self.sender = get_sender_text(str(self.sender))
    
    def get_features(self):
        if self.category == 'none':
            self.features = self.body+' '+self.subject+' '+self.to+' '+self.sender
        else:
            self.features = self.body+' '+self.subject+' '+self.to+' '+self.sender+' '+self.category
            
# outlook-specific functions, classes, etc
# functions, text templates, and category tokens
def build_outlook_path(predicted_path):
    f='.Folders[{}]'*len(predicted_path.split(';'))
    l=[]
    for a,b in list(zip(f.split('.')[1:],predicted_path.split(';'))):
        b="'"+b+"'"
        c=a.format(b)
        l.append(c)
    return 'socs_main_folder.'+'.'.join(l)

def format_outlook_path(path_string,rcvd):
    if "yy" in path_string.lower():
        date_parts = [rcvd.strftime("%m"),rcvd.strftime("%B"),rcvd.strftime("%Y")]
        if "TEAM" in path_string:
            for r in (("dd","*"+date_parts[0]),('Mmmm',date_parts[1]),("yyyy",date_parts[2])):
                path_string = path_string.replace(*r)
            return path_string
        else:
            return path_string.replace("YYYY","*"+date_parts[2])
    else:
        return path_string
    
# template text
c_ai_text = "C_AI[READING]:{e}\nC_AI[CATEGORIZED]:{p}\n------"
f_ai_text = "F_AI[READING]:{e}\nF_AI[LOCATION]:{f}\n------"
f_ai_text2 = "F_AI[READING]:{e}\nF_AI[LOCATION]:{f}\nF_AI[M]:{m}\nF_AI[M:ATT]:{a}\n------"

class email_obj(object):
    def __init__(self,email):
        self.body = email.Body
        self.subject = email.Subject
        try:
            self.sender = email.Sender.Address
        except:
            self.sender = "None"
        try:
            self.to = email.To
        except:
            self.to = "None"
        self.received = email.ReceivedTime
        self.attcount = email.Attachments.Count
        if not email.Categories:
            self.category = 'none'
        else:
            self.category = email.Categories
            if self.category == "File":
                self.prot = re.search('(\d{2}\-\d{4})',self.subject)[0]
                self.yr_part = self.prot.split("-")[0]
                self.outlook = "Inbox;{p};Completed".format(p=self.prot)
                self.mdrive = "custom/path/to/study/{}/Emails/".format(y=self.yr_part,p=self.prot)
                self.attpath = "custom/path/to/study/{}/Emails/Attachments/".format(y=self.yr_part,p=self.prot)
        
    def clean_attributes(self):
        x = parse_email(clean_text_content(prepare_body(str(self.body))))['body']
        x = re.sub('\w+:([^.]*)*','',x)
        x = re.sub('[^a-zA-Z]+', ' ', x)
        self.body = x
        self.sender = get_sender_text(str(self.sender))
    
    def get_features(self):
        if self.category == 'none':
            self.features = self.body+' '+self.subject+' '+self.to+' '+self.sender
        else:
            self.features = self.body+' '+self.subject+' '+self.to+' '+self.sender+' '+self.category
            
