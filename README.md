# Welcome to my Project Portfolio
This repository contains example implementations of programmatic solutions I have developed and deployed in my professional career for stakeholders in the medical/clinical industry. My portfolio includes a mix of traditional automation methods and neural network powered "digital worker" automation. While I am happy to program anything required of me, I have so far focused on automated solutions because they save organizations time and resources.

I am interested in leveraging artificial intelligence combined with programmatic automation in a variety of ways ranging from practical business applications to unique scientific projects. For example, my latest natural language processing project shortened turnaround time and improved quality of email-related administrative tasks typically done by humans. Below is a list of projects that I have built and deployed to solve challenges for colleagues met in the course of my career.

________

# Content
### Contact Info Comparison App (v1.0 1/10/19)
[Contact Info Comparison App](https://github.com/acrovitic/python_portfolio/tree/master/system_interfacing/contact_comparison_app "Contact Info Comparison App")

GUI-based application that identifies inconsistencies in research project team contact information between a document in a shared drive and data in a content management website.

### Deep Neural Network for Research Operations (v1.0 12/27/18)
[Deep Neural Network](https://github.com/acrovitic/python_portfolio/tree/master/deep_neural_network "Deep Neural Network")

v1.0 capabilities are as followes:

* Categorize email (1 of 37 categories)
* Move email to Outlook folder (1 of 150+ folders)
* Move email to hard drive folder (1 of 150+ folders)
* Save email attachment to relevant location (1 of 150+ folders)
* Report on actions as they occur per email

Many tasks in research operations are better managed by artificial intelligence. To determine how many tasks an A.I. could handle, I developed two Deep Neural Networks (DNN) to manage basic functions of a shared clinops inbox. Both DNNs will be continually improved upon, with the ultimate goal being a single robust DNN that can remain operational 24/7, continually monitoring emergent events in a research project, and react accordingly through a diverse array of actions normally taken by humans.

### Recursive Functions for Outlook Folder Looping (v1.0 12/27/18)
[Depth Detector and Pathfinder](https://github.com/acrovitic/python_portfolio/tree/master/outlook_automation/folder_recursion "Loop through Outlook folders and all subfolders with the Outlook Depth Detector and Outlook Pathfinder functions")

I once read that there were no easy methods to recursively travel through Outlook folders. To address this issue, I created two functions; `outlook_depth_detector` and `outlook_depth_pathfinder`. These functions can be modified to suit multiple types of recursive activities in Outlook.

### Study Status Manager (v1.0 12/27/18)
[Study Status Manager](https://github.com/acrovitic/python_portfolio/tree/master/clinical_operations_automation/study_status_management "Study Status Manager")

Time to complete task BEFORE script | Time to complete task AFTER script 
--- | --- 
45 minutes | 25-35 seconds 

A regularly generated report required data points from five seperate reports, each in a different location. Any data point shared by two or more reports needed to match across the board. Originally, an employee would need to review all records in these reports on a weekly basis to find discrepancies, notify the team member assigned to the discrepant study of the the issue, request they address the issue, and track their progress to ensure they address the issue.

This script completed this task automatically, including emailing the appropriate team members with a formatted list of task and study number to address.

### Research Document Transmission (v1.0 12/27/18)
[Document Distributor](https://github.com/acrovitic/python_portfolio/tree/master/outlook_automation/clinical_document_distributor "Document Distributor")

Time to complete task BEFORE script | Time to complete task AFTER script 
--- | --- 
7-10 minutes per document | 6-7 seconds per document 

We recieve updated Protocol documents from multiple studies at irregular times. These documents must be posted online to two 
different websites, placed in specific folders in a shared drive (while moving off their previous versions), and sending
them to key stakeholders to update critical clinical documents. This must be done for each updated Protocol received.

This script handles all of the above described tasks in mere seconds.

### Automatic Adverse Event Reporting (v1.0 12/27/18)
[SAE Automater.](https://github.com/acrovitic/python_portfolio/tree/master/clinical_operations_automation/sae_automation "SAE Automater") 

Time to complete task BEFORE script | Time to complete task AFTER script 
--- | --- 
12-15 minutes per report | 10 seconds per report

This script collects all matching reports and notifications for an adverse event, combines them into a formatted email for distribution, checks for previous emails regarding the particular event, and attaches them if the intended recipient has not responded yet.

### Independent Monitoring Committee Candidate Invite Generator (v2.0 12/27/18)
[Invite Generator.](https://github.com/acrovitic/python_portfolio/tree/master/clinical_operations_automation/dsmb_candidate_invite_generator "Invite Generator") 

Time to complete task BEFORE script | Time to complete task AFTER script 
--- | --- 
30-35 minutes per invite | 6 seconds per invite

Example implementation to automate committee invitation package generation. Fills out templates, writes/formats emails, attaches completed templated documents to outlook email.

v2.0 update: 

* Replaced most pandas-based data transformations with python native approaches. 
* Wrapped everything within one class.
* Improved command-prompt instructions/interface.

### Meeting Summary Generator
[Meeting Summary/Action Item Generator.](https://github.com/acrovitic/python_portfolio/tree/master/outlook_automation/actionitems "Meeting Summary/Action Item Generator") Creates formatted outlook email with action items from a team meeting.

### Miscellaneous
[Miscellaneous Scripts.](https://github.com/acrovitic/python_portfolio/tree/master/Miscellaneous "Random Bag'o'Fun") Miscellaneous scripts and programs made to help colleagues quickly complete various smaller tasks.

### Automated Document Migration 
#### Data cleaning and uploading (Migration Part 2)
[Document Migration Part II: document cleaner/uploader.](https://github.com/acrovitic/python_portfolio/tree/master/system_interfacing/mass_uploading "Document Migration Part II: document cleaner/uploader") Second step in migrating 1000 documents from an old clinical document repository to a new one. Previous anticipated time to completion was 4 weeks. This script completed the task in 2 hours.

#### Downloading (Migration Part 1)
[Document Migration Part I: Batch document downloader/handler.](https://github.com/acrovitic/python_portfolio/tree/master/system_interfacing/mass_downloading "Document Migration Part I: Batch document downloader/handler") First step in migrating 1000 documents from an old medical document repository to a new one. Previous anticipated time to completion was 1 week. This script completed the task in roughly 1 minute 30 seconds.

### OBSOLETE (replaced by DNN) - Outlook Shared Inbox Autofiler
[Inbox Filer.](https://github.com/acrovitic/python_portfolio/tree/master/outlook_automation/shared_inbox_filer "Inbox Filer") 

Time to complete task BEFORE script | Time to complete task AFTER script 
--- | --- 
15-40 minutes (depending on number of emails received) | < 60 seconds  

Files emails, saves attachments and emails to appropriate folder on hard drive and outlook. 
