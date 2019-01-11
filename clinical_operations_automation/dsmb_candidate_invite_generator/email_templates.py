noprior = (r"""
<body>
<font face = "Times New Roman" size="3">
<p><strong>Protocol:</strong> {Protocol_Number}<br>
<strong>Title:</strong> {Protocol_Full_Title}</p>
<p>Dear Dr. {Last_Name},</p>
<p>We would greatly appreciate your participation on this {Committee} as it is an important part of the safety oversight 
process within MRA. Participation generally entails teleconference meetings to review safety data at intervals 
specified by the {Committee_Type} Charter. These teleconferences will include study team members including the Principal 
Investigator and MRA staff.</p>
<p>Attached is the Welcome Letter regarding this {Committee_Type} as well as a Contact Information Form referenced in the letter. 
If you are willing to participate, please complete the necessary documentation and respond to me directly at 
<a href="mailto:CSS@medicalresearch.com">CSS@medicalresearch.com</a> by {due_date_formatted}.</p>
<p>If you have any questions, or require clarifications related to your participation, please contact me directly 
at CSS@medicalresearch.com or 301-222-3333.</p>
<p>Sincerely,<br>
John Smith, MS<br>
Clinical Safety Support<br>
Example Corporation<br>
1234 Main Street, Suite 9001<br>
Bethesda, MD 20817<br>
Phone: 301-222-3333<br>
Fax: 301-222-4444<br>
<a href="mailto:jsmith@example.com">jsmith@example.com</a><br>
<a href="http://www.example.com">www.example.com</a></p>
</font></body>
""")

prior = (r"""
<body>
<font face = "Times New Roman" size="3">
<p><strong>Protocol:</strong> {Protocol_Number}<br>
<strong>Title:</strong> {Protocol_Full_Title}</p>
<p>Dear Dr. {Last_Name},</p>
<p>On behalf of MRA, we would like to thank you for expressing interest in serving as a Member of the {Committee} for 
Protocol {Protocol_Number}.</p>
<p>Attached is the Welcome Letter regarding this {Committee_Type} as well as a Contact Information Form referenced in the letter. 
If you are willing to participate, please complete the necessary documentation and respond to me directly at 
<a href="mailto:CSS@medicalresearch.com">CSS@medicalresearch.com</a> by {due_date_formatted}.</p>
<p>If you have any questions, or require clarifications related to your participation, please contact me directly 
at CSS@medicalresearch.com or 301-222-3333.</p>
<p>Sincerely,<br>
John Smith, MS<br>
Clinical Safety Support<br>
Example Corporation<br>
1234 Main Street, Suite 9001<br>
Bethesda, MD 20817<br>
Phone: 301-222-3333<br>
Fax: 301-222-4444<br>
<a href="mailto:jsmith@example.com">jsmith@example.com</a><br>
<a href="http://www.example.com">www.example.com</a></p>
</font></body>
""")
