# Automation in clinical research
Portfolio of scripts I have developed and deployed to automate administrative and operational tasks at my clinical research-focused job. 
https://www.linkedin.com/in/vpoonai/

# Content
### Deep Neural Network for Clinical Research (v1.0 12/27/18)
[Deep Neural Network](https://github.com/vpoonai/Office-Automation/tree/master/deep_neural_network "Deep Neural Network")

v1.0 capabilities are as followes:

Categorize emails (1 of 37 categories)

Move emails to Outlook and hard drive folders (1 of 150+ folders)

Many tasks in clinical research operations can be better managed by artificial intelligence. To determine how many ClinOps tasks an A.I. could handle, I have developed two Deep Neural Networks (DNN) to manage basic functions of a shared safety oversight inbox. This DNN will be continually improved to include more and more tasks. The ultimate goal is to create a robust DNN that outperforms human ClinOps and clinical research administrators workers by being active and responsive 24/7.

### Recursive Functions for Outlook Folder Looping (v1.0 12/27/18)
[Depth Detector and Pathfinder](https://github.com/vpoonai/Office-Automation/tree/master/outlook_automation/folder_recursion "Loop through Outlook folders and all subfolders with the Outlook Depth Detector and Outlook Pathfinder functions")

I read that there was no easy way to recursively travel through Outlook folders. To address this issue, I created the two functions; outlook_depth_detector and outlook_depth_pathfinder. 

These functions can likely be modified to suit multiple types of recursive activities in Outlook.

### Study Status Manager (v1.0 12/27/18)
[Study Status Manager](https://github.com/vpoonai/Office-Automation/tree/master/clinical_operations_automation/study_status_management "Study Status Manager")

Time to complete task (before script): 45 minutes

Time to complete task (after script): 25-35 seconds

Study datapoints are recorded in multiple locations. These locations must match one another. I am tasked with reviewing all records on a weekly basis, identifying discrepancies, notifying the team member assigned to the discrepant study of the the issue, requesting they address the issue, and tracking their progress to ensure they follow through.

This script completed this task automatically, including emailing the appropriate team members.

### Clinical Document Distributor (v1.0 12/27/18)
[Document Distributor](https://github.com/vpoonai/Office-Automation/tree/master/outlook_automation/clinical_document_distributor "Document Distributor")

Time to complete task (before script): 5-10 minutes per document

Time to complete task (after script): 6-7 seconds per document

We recieve updated Protocol documents from multiple studies at various times. These documents must be posted online to two 
different websites, placed in specific folders in a shared drive (while moving off their previous versions), and sending
them to key stakeholders to update critical clinical documents. This must be done for each updated Protocol received.

This script handles all of the above described tasks, completing the overall task in a few seconds.

### Serious Adverse Event (SAE) Reporting Automation (v1.0 12/27/18)
[SAE Automater.](https://github.com/vpoonai/Office-Automation/tree/master/clinical_operations_automation/sae_automation "SAE Automater") 

Time to complete task (before script): 10-15 minutes per SAE pair

Time to complete task (after script): 10-12 seconds per SAE pair

SAE email subject lines are uniquely formatted at my office. This script utilizes SAE email subject lines as a 
unique identifier to match notifications and full narrative reports.
Once matched, the script searches for additional emails related to the specific SAE to determine whether the content should be sent along with a reminder message, or if another context should be used. 
Finally, after the search is completed, the content is cleaned and placed into a pre-formatted email template, which in turn is placed into the body of a new Outlook email instance.
### DSMB/SMC Candidate Invite Generator (v2.0 12/27/18)
[Invite Generator.](https://github.com/vpoonai/Office-Automation/tree/master/clinical_operations_automation/dsmb_candidate_invite_generator "Invite Generator") 

Time to complete task (before script): 30-35 minutes per invitation

Time to complete task (after script): 6 seconds per invitation

Example implementation to automate safety committee invitation package generation. Fills out templates, writes/formats emails, attaches completed templated documents to outlook email.

v2 update: 

Replaced most pandas-based data transformations with python native approaches. 
Wrapped everything within one class.
Improved command-prompt instructions/interface.
### Automated Document Migration (Part 2)
[Document Migration Part II: document cleaner/uploader.](https://github.com/vpoonai/Office-Automation/tree/master/system_interfacing/mass_uploading "Document Migration Part II: document cleaner/uploader") Second step in migrating 1000 documents from an old clinical document repository to a new one. Supervisor expected task to be completed in 4 weeks. This script I wrote allowed me to complete the task in 3 hours.
### Automated Document Migration (Part 1)
[Document Migration Part I: Batch document downloader/handler.](https://github.com/vpoonai/Office-Automation/tree/master/system_interfacing/mass_downloading "Document Migration Part I: Batch document downloader/handler") First step in 
migrating 1000 documents from an old clinical document repository to a new one. Supervisor expected task to be completed in 1 week. This script I wrote allowed me to complete the task in roughly 1 minute 30 seconds.
### Outlook Shared Inbox Autofiler
[Inbox Filer.](https://github.com/vpoonai/Office-Automation/tree/master/outlook_automation/shared_inbox_filer "Inbox Filer") 

Time to complete task (before script): 15-40 minutes (depending on previous days volume of emails)

Time to complete task (after script): >=1 minute (depending on previous days volume of emails)

Files emails saves attachments and emails to appropriate clinical trial folder on hard drive and outlook. Coming Soon: Selenium script to upload attachments
to required document repositories (80% complete - final testing phase).
### Meeting Summary Generator
[Meeting Summary/Action Item Generator.](https://github.com/vpoonai/Office-Automation/tree/master/outlook_automation/actionitems "Meeting Summary/Action Item Generator") Creates formatted outlook email with action items from a team meeting.
### Miscellaneous
[Miscellaneous Scripts.](https://github.com/vpoonai/Office-Automation/tree/master/Miscellaneous "Random Bag'o'Fun") Miscellaneous scripts and programs made to help colleagues quickly complete various tasks.
