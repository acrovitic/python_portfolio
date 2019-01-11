# functions

def remove_dupes(l, k):
    seen = {} 
    for d in l:
        v=d[k]
        if v not in seen:
            seen[v] = d
    return seen

def group_by(dictionary_list,getter,finder): # 'protocol', 'path' for example
    key = operator.itemgetter(getter)
    b = [{getter: x, finder: [str(d[finder]) for d in y]} 
         for x, y in itertools.groupby(sorted(dictionary_list, key=key), key=key)]
    for d1 in dictionary_list:
        for d2 in b:
            if d1[getter] == d2[getter]:
                d1[finder+"_"] = d2[finder] 

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

def get_title(the_iterable, condition = lambda x: True):
    for n,i in enumerate(the_iterable):
        if condition(i):
            return the_iterable[n+1]

def find_item(the_iterable, condition = lambda x: True):
    for n,i in enumerate(the_iterable):
        if condition(i):
            return i

def get_dates(string):
    string=str(string)
    month=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    match=re.search('(\d{1,2}\/\d{1,2}\/\d{2,4})|(\d{1,2}\ +\w{3,9}\ +\d{2,4})|(\d{1,2}\w{3}\d{2})|(\w{3,9}\ +\d{1,2},\s+\d{4})',string)
    if match:
        date=match[0]
        try:
            d = parse(date).strftime('%d%b%y')
            return d
        except:
            pass

def get_version(string):
    match = re.search('(\d{1,2}\.\d{1})',string)
    if match:
        return "v"+match[0]
    else:
        return "version"
    
def get_icon_info(acronym): 
    icon_info = {
        "KM":["Kandace","email1@example.com"],
        "ER":["Ellen","email2@example.com"],
        "HT":["Harry","email3@example.com"],
        "SB":["Sally","email4@example.com"],
        "VK":["Vish","email5@example.com"],
        "LR":["Larry","email6@example.com"]
    }[acronym]
    return icon_info

# get assigned icon personnel
def get_icon_personnel(prot_string):
    ipath = "path/to/Protocol Assignments/"
    list_of_files = glob.glob(ipath + 'Staff Protocol Assignments*')
    latest_file = max(list_of_files, key=os.path.getctime)
    df2 = pd.read_excel(latest_file,skiprows=1)
    mail_dict = df2.to_dict('records')
    if not any(d['Protocol #'] == prot_string for d in mail_dict):
        return ["Karey","karey.smith@ateamcompany.com"]
    else:
        for d in mail_dict:
            if d['Protocol #'] == prot_string:
                return get_icon_info(d['TEAM'])

def download_protocol(file_url):
    site2 = "www.website1.com/signin/auth"
    login = "www.website1.com/signin"
    site = "www.website1.com/default"
    file_path_template = "path/to/studies/Essential Documents/{f}"
    username = username
    password = password
    ua = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.110 Safari/537.36'
    output_file_name = file_url.rsplit("/",1)[1]
    file_year = re.search('(\d{2})\-\d{4}',output_file_name)[1]
    file_protocol = re.search('(\d{2}\-\d{4})',output_file_name)[1]
    output_file_path = file_path_template.format(y=file_year,p=file_protocol,f=output_file_name)
    data = {'txtUserName':username,'txtPassword':password}
    s = requests.Session()
    s.headers['User-Agent'] = ua
    r1 = s.get(site2)
    r1 = s.get(site2)
    soup = BeautifulSoup(r1.text,'html.parser')
    for elem in soup.form.findAll('input'):
        try:
            if 'value' in elem.attrs.keys():
                data[elem['name']] = elem['value']
            else:
                data[elem['name']] = data[elem['name']]
        except:
            continue
    data["btnSignIn.x"] = 25
    data["btnSignIn.y"] = 3
    data['RememberMe'] = 'on'
    r2 = s.post(site2,data=data)
    r2 = s.get(file_url)
    if r2.status_code == 200:
        with open(output_file_path, 'wb') as f:
            for chunk in r2.iter_content():
                f.write(chunk)
        f.close()
        return output_file_path
    else:
        return "Response Code: ", r2.status_code, "\nResponding URL: ",r2.url
