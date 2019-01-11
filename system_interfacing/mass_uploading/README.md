### Task: Upload 1000 documents from old web-based document management system to new system, with all metadata fields filled out and documents associated to meetings as necessary.
### Supervisor estimated task completion time: 4 weeks
### Script dev time: 6 hours
### Actual task completion time with script: 3 hours 15 minutes

### Narrative
After downloading 1000 documents from the to-be-retired document repository, I was tasked with upload all downloaded documents to the new
repository. Uploading included entering metadata into form fields in the new website, associating documents with meetings they 
were used in, and filling in gaps in metadata due to differences in repository formatting (e.g. meeting dates were coded in very 
different formats. My supervisor believed that this could only be addressed comparing by hand each documents meeting data to the new 
website's existing meeting date data for each document).

I fully automated this task by using a combination of custom date matching functions and date cleaning classes.
A dictionary was generated for each document to hold its associated metadata.
Additional metadata was written to each dictionary in preparation for the final step of this document migration project.
Data cleaning and formatting was handled through a custom 'document' class.
Dictionaries were fed to Selenium to upload each document.
An excel report containing all metadata was generated after all documents were uploaded.
