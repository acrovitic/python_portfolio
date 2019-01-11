# email templates
email_subject = "Safety Oversight, Protocol {protocol}, New Protocol({version})"
email_template = (r"""
<body>
<font face = "Times New Roman" size="11pt">
<strong>Protocol:</strong> {protocol}<br>
<strong>Title:</strong> {title}<br>
<p>Dear {TEAM[0]}</p>
<p>
Please find attached the Protocol {version} {date} of Protocol {protocol}{msg_part}
</p>

<p>If you have any questions, or require clarifications , please contact me directly 
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
