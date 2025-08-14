import re
import numpy as np
from dateutil.parser import parse

def get_dates(string):
    string = str(string)
    month = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
    match = re.search(
        '(\d{1,2}\/\d{1,2}\/\d{2,4})|(\d{1,2}\ +\w{3,9}\ +\d{2,4})|(\d{1,2}\w{3}\d{2})|(\w{3,9}\ +\d{1,2},\s+\d{4})',
        string)
    mtgtypes = ['data review', 'org', 'ad hoc', 'ad-hoc', 'e-rev', 'e rev', 'elec', 'drm', 'follow-up', 'follow up',
                'meeting']
    if not any(i in string.lower() for i in mtgtypes):
        return np.nan
    else:
        if match:
            date = match[0]
            if len(date) == 7 and re.match('(\d{1,2}\w{3}\d{2})', date):
                if any(i in date.lower() for i in month):
                    return parse(date).strftime('%m/%d/%Y')
                else:
                    return np.nan
            else:
                return parse(date).strftime('%m/%d/%Y')
        else:
            return np.nan


def date_matcher(string):
    string = str(string).replace("\ISM", "")
    if 'H5N1' in string:
        string = string.replace('ABC-', '').replace('-EFG', '')
        return string
    if ":" in string:
        string_date_parts = string.rsplit('-', 1)
        string_date_parts[1] = parse(''.join(string_date_parts[1].rsplit(' ', 1)[:1])).strftime('%m/%d/%Y')
        return ' '.join(string_date_parts)
    else:
        return string.replace("DSMB-", "DSMB 0")


def get_access(string):
    string = string.lower()
    if "close" in string or "unblind" in string:
        return "Closed"
    else:
        return "Open"


def get_version(file_name):
    match1 = re.search("v(\d{1,2}\.\d{1})", file_name)
    match2 = re.search("(\d{1,2}\.\d{1})", file_name)
    if not match1 and not match2:
        return "1.0"  # for batch 2 ONLY. change later
    elif not match1:
        return match2[1]
    else:
        return match1[1]