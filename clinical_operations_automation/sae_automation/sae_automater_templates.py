# sae_initial_email template
sae_initial_email = (r"""
<body>
<font face = "Times New Roman" size="11pt">
<strong>Protocol:</strong> {prot}<br>
<strong>Title:</strong> {title}<br>
<p>Dear Dr. X</p>
<p>
The site should have already provided you, as the ISM, the SAE information for Protocol {prot}.  
In addition, as per the Charter, you are being notified of this event. The <em> initial </em> site notification 
along with the Narrative Summary, which have not yet been formally reviewed by the Medical Monitor, 
are provided below.  
</p>
<p>
Please "Reply All" by <strong><u>{day}</u></strong> and:
<ol type="1">
  <li>Provide subject-specific information regarding this event for the other Committee members to review.</li>
  <li>Indicate, based on your assessment, if an Ad hoc meeting is necessary.</li>
  <li>Indicate if you will be requesting additional information from the site.</li>
</ol>  
</p>
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
<p><hr></hr></p>
<body>
<font face = "Times New Roman" size="11pt">
<strong>Protocol:</strong> {prot}<br>
<strong>Title:</strong> {title}<br><br>
</font></body>
{notification}
<body>
<font face = "Times New Roman" size="11pt">
</font></body>
{report}
""")

# sae_followup_email template
sae_followup_email = (r"""
<body>
<font face = "Times New Roman" size="11pt">
<strong>Protocol:</strong> {prot}<br>
<strong>Title:</strong> {title}<br>
<p>Dear Dr. X</p>
<p>
You previously reviewed the SAE in Subject {subject_id}, <em>{condition}</em>, and did not request an 
Ad hoc Meeting. The <em>{followup_ord}</em> follow-up site {sp1} along with the {sp2}, 
which {sp3} not yet been formally reviewed by the Medical Monitor, {sp4} provided below for your information.  
</p>
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
<p><hr></hr></p>
<body>
<font face = "Times New Roman" size="11pt">
<strong>Protocol:</strong> {prot}<br>
<strong>Title:</strong> {title}<br><br>
</font></body>
{notification}
<body>
<font face = "Times New Roman" size="11pt">
</font></body>
{report}
""")

# sae_initial_fu_reminder
# if response to initial not found in xx-xxxx->Completed folder
# iterate from current f/u number down to 1, where 1 is initial notice
sae_init_fu_email = (r"""
<body>
<font face = "Times New Roman" size="11pt">
<strong>Protocol:</strong> {prot}<br>
<strong>Title:</strong> {title}<br>
<p>Dear Dr. X</p>
<p>
We previously provided you with the {fuhist} site {sp1} along with {sp2}, which has not yet been formally reviewed
by the Medical Monitor, regarding the SAE in Subject {subjectid}, {condition}, but have not yet received your
response (see attached messages).  
</p>
<p>
In addition, the <em>{followup_ord}</em> follow-up site {sp1} and {sp2}, which {sp3} not yet been formally reviewed by 
the Medical Monitor, {sp4} provided below.  
</p>
<p>
Please "Reply All" by <strong><u>{day}</u></strong> and:
<ol type="1">
  <li>Provide subject-specific information regarding this event for the other Committee members to review.</li>
  <li>Indicate, based on your assessment, if an Ad hoc meeting is necessary.</li>
  <li>Indicate if you will be requesting additional information from the site.</li>
</ol>  
</p>
</p>
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
<p><hr></hr></p>
<body>
<font face = "Times New Roman" size="11pt">
<strong>Protocol:</strong> {prot}<br>
<strong>Title:</strong> {title}<br><br>
</font></body>
{notification}
<body>
<font face = "Times New Roman" size="11pt">
</font></body>
{report}
""")
