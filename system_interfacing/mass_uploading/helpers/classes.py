from selenium.webdriver.common.action_chains import ActionChains
import time

from utils.utils import *

# custom class to add sleep action to Selenium action chains
class Actions(ActionChains):
    def wait(self, time_s: float):
        self._actions.append(lambda: time.sleep(time_s))
        return self


class document(object):
    protnum_associations = {'old assoc': 'new assoc'}
    old_path = "path/to/files/old"
    new_path = 'path/to/files/new'
    switch_dict = {'old': 'new'}

    def __init__(self, dictionary):
        for k, v in dictionary.items():
            setattr(self, k, v)
        self.BRANCH = self.BRANCH.split('#')[1]
        self.MEETING = self.MEETING.strip()
        self.mtg_date = get_dates(self.MEETING)
        self.version = get_version(self.Name)
        self.access = get_access(self.Name)

    def get_recipients(self, cpm_data):
        recipients = ["placeholder"]
        for d in cpm_data:
            if d['Branch'] == self.cpm_branch:
                recipients.append(d['CPM'])
        self.recipients = recipients

    def get_uploading_name(self):
        if self.ProtocolName in self.Name:
            self.uploading_name = self.Name
        else:
            self.uploading_name = f"{self.ProtocolName}_{self.Name}"

    def associate_protocol(self):
        for k, v in document.protnum_associations.items():
            if self.ProtocolName in v:
                self.ProtocolGroup = k
            else:
                self.ProtocolGroup = self.ProtocolName

    def get_mtg_data(self, meeting_dict):
        if self.mtg_date is np.nan:
            self.mtg_name = 'select'
            self.mtg_id = 0
        else:
            self.mtg_check1 = f"{self.BRANCH}-{self.ProtocolName}-{self.cmt_type} {self.mtg_date}"
            for d in meeting_dict:
                if date_matcher(self.mtg_check1) == date_matcher(d['MeetingName']):
                    self.mtg_name = d['MeetingName']
                    self.mtg_id = d['MeetingID']
                    break
                else:
                    if self.ProtocolName in document.protnum_associations:
                        for i in document.protnum_associations[self.ProtocolName]:
                            self.mtg_check = f"{self.BRANCH}-{i}-{self.cmt_type} {self.mtg_date}"
                            if date_matcher(self.mtg_check) == date_matcher(d['MeetingName']):
                                self.mtg_name = d['MeetingName']
                                self.mtg_id = d['MeetingID']
                                break
                    else:
                        self.mtg_name = 'select'
                        self.mtg_id = 0

    def get_mtg_button_type(self):
        if "Minute" in self.DOC_TYPE or "Recomm" in self.DOC_TYPE:
            self.mtg_button_type = 'Post-Meeting'
        else:
            self.mtg_button_type = 'Meeting'

    # functions to remove special characters and rename paths if necessary
    def clean_uploading_name(self):
        exclude = "!@#$%^&*()[]{};:,/<>?\|`~'=+"
        if any(i in self.uploading_name for i in exclude):
            self.uploading_name = self.uploading_name.translate({ord(c): "" for c in exclude})
        else:
            pass

    def get_location(self):
        self.location = f"{document.old_path}/{self.PROTOCOL}/{self.MEETING}/{self.Name}"

    def get_clean_location(self):
        self.clean_location = f"{document.old_path}/{self.PROTOCOL}/{self.MEETING}/{self.uploading_name}"

    def get_destination(self):
        yr_part = re.search('(\d{2})\-\d{3,4}', self.PROTOCOL)
        if yr_part:
            self.destination = f"{document.new_path}/{yr_part[1]}{self.PROTOCOL}"
        else:
            if self.PROTOCOL in document.switch_dict.keys():
                self.destination = f"{document.new_path}/{yr_part[1]}{self.PROTOCOL}"

    # functions to bundle protocol associated/mtg associated docs
    def get_posting_association(self):
        if 'email' in self.Name or " em " in self.Name:
            self.association = 'none'
        else:
            if int(self.mtg_id) == 0:
                self.association = 'protocol'
            if int(self.mtg_id) > 0:
                self.association = 'meeting'