### Task: Download 1000 files
### Supervisor estimated task completion time: 1 week
### Script dev time: 1 hour
### Actual task completion time with script: 1 minute 27 seconds.

### Narrative
I was tasked with downloading 1000 files from a clinical document repository that would soon be retired. My supervisor estimated the task 
would take a week and provided me with a sheet listing the documents along with their meta data.
When inspecting the document repository's URL structure for each document, I discovered that the meta data in the sheet I was given could be
used to construct each file's direct download URL. I then built this script to automate file download and shared drive placement. 
Pandas was used to assemble the paths for the download URL, folder to download the file to, and new custom file paths for each file. Each
custom file path mirrored the repository's folder structure, allowing team members to quickly find downloaded files.
Selenium was used to interact with the repository via Chrome and download files.
Standard library os was used to create directories mirroring downloaded file locations in the repository and move files to these
new directories.
