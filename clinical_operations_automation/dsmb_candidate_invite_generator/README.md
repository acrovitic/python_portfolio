# Executive Summary
Example implementation to automate safety committee invitation package generation. Fills out templates, writes/formats emails, attaches
completed templated documents to outlook email.

# Full Description
Some clinical trials require safety oversight committees, such as Data and Safety Monitoring Boards (DSMB) or Safety Monitoring 
Committees (SMC).

To form these committees, subject matter experts are sent invitation packages requesting their service. Packages include Welcome Letters,
Contact Informtion sheets, Conflict of Interest (COI) Disclosure forms, and a summary document of the clinical trial in question.

Creating these packages manually can be time intensive and ties up resources needed elsewhere (20-30 minutes per invite). 
This folder contains an example automated invitation generator (1-2 seconds per invite) for safety oversight committee candidates 
using pandas, mailmerge, and other methods. 

Using data sources generated from a centralized system, contact information and trial information are retrieved and placed into templated
documents and emails for as many candidates as needed. Data sources in this case are Excel spreadsheets, as my team is not permitted
direct access to the data base. Luckily, Python behaves very well with Excel spreadsheets, so I was able to work around not having
DB access.

Since I work in a corporate environment, I built this script with Microsoft Office products (specifically, Word, Excel, and Outlook) in mind. 
If enough people request that this script be adapted to other programs, I will try to do so.

When run from a cmd prompt in the folder containing main script (invite_gen_main.py), the following process occurs:

1-User is prompted to enter "clinical trial number,first name,last name,role,noprior/prior-start/mid."

1a-Subject matter experts can be invited under a number of different scenarios requiring unique language. To account for this, 
   the "noprior/prior-start/mid" identifies which type of templated communication the script will populate.
   
1b-noprior=intended recipient has not yet been contacted by someone (e.g. a federal employee) about this invite ahead of it being sent.

1c-prior=intended recipient has been contacted by someone (e.g. a federal employee) about this invite ahead of it being sent.

1d-start=intended recipient is being invited at the beginning of the clinical trial.

1e-mid=intended recipient is being invited after the clinical trial has started.

2-User can enter as many contacts as they desire. When finished, they may press "Enter" on an empty prompt.

3-Data source sheets are parsed for the appropriate information. The final product is placed into a dictionary to allow for the fastest
  package generation possible.
  
4-Script will select and populate Word and email templates based on their file name.

4a-Word templates are populated by pre-placed mail merge tags and dictionary data.

4b-Email templates are populated by string format tags and dictionary data.

5-Outlook email is generated and fully populated with email address, attachments, and email body fully formatted and ready to send.

Example output is located in the "output" folder.

NOTE: The main contact information data source (welcomepackage_report) is empty in the interest of confidentiality. However, all columns
      are correctly named and can be used if one were to fill out the empty sheet with their own data as needed.
      Clinical trial information in the protocol_title sheet is not empty, but is openly available through the appropriate government websites,
      and therefore have no issue with confidentiality.
