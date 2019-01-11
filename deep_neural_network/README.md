### Transforming Clinical Research Through Artificial Intelligence
Many tasks in clinical research operations can be better managed by artificial intelligence. To determine how many ClinOps tasks an 
A.I. could handle, I have developed two Deep Neural Networks (DNN) to manage basic functions of a shared safety oversight inbox.

The DNNs assigned one of 37 categories to uncategorized emails and moves emails to the appropriate Outlook and hard drive folder (one of
over 150 possible folders). After training, the DNNs performs with 97.3% accuracy and is currently deployed.

The next step is to combined both DNNs into one and expand its capabilities to differentiate between emails requesting some action be taken, emails providing information that must be used to update some report, or emails unique to a team member's assigned study that should be forwarded to 
said team member. 

My hope is that this method of classification will allow the DNN to execute scripts based on the email's content, 
much like a human would. For example, the DNN would recognize a new protocol version email and run the clinical_document_distributor
script automatically. This would improve turnaround time and efficiency in clinical trials at a level that humans cannot operate at.
