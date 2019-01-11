def mtg_switch(string):
    return {"Data Review":"DRM",
              "Organizational":"ORG",
              "E-Review":"E-Rev","Ad-Hoc":"Ad Hoc"}[string]

def get_email_address(string):
    return {"DG":"email1@example.com",
            "AN":"email2@example.com",
            "HM":"email3@example.com",
            "MP":"email4@example.com",
            "MH":"email5@example.com",
            "VP":"email6@example.com",
            "TT":"email7@example.com",
            "TJ":"email8@example.com"}[string]

def get_name(string):
    return {"DG":"Diana",
            "AN":"Anthony",
            "HM":"Hanna",
            "MP":"Megan",
            "MH":"Marcy",
            "VP":"Vic",
            "TT":"Tina",
            "TJ":"Ted"}[string]

def group_by(dictionary_list,getter,finder): # 'protocol', 'path' for example
    key = operator.itemgetter(getter)
    b = [{getter: x, finder: [str(d[finder]) for d in y]} 
         for x, y in itertools.groupby(sorted(dictionary_list, key=key), key=key)]
    for d1 in dictionary_list:
        for d2 in b:
            if d1[getter] == d2[getter]:
                d1[finder+"_"] = d2[finder] # adds finder [actions] into list for finders sharing getter [person.
                # dedup later to have one meeting id with multiple doc ids as needed.

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

def intersperse(lst, item):
    result = [item] * (len(lst) * 2 - 1)
    result[0::2] = lst
    return result
