import calendar
import pandas as pd
import datetime as dt
import email_templates
import win32com.client as win32
from datetime import date, timedelta

# Set date variables for action item table, email body, and subject line
today = dt.date.today()
next_week = today + dt.timedelta(days=7)
d = today.strftime('%d%b%y')

# Space to enter action items. Format below (action item.:Initials|) must be followed for table generation.
action_items = '''
Train new hire on internal document manager system.:ES|
Follow up with IT regarding connectivity issues.:MT|
Update automated data cleaning script and provide testing results.: VP
'''

# Create action item table. Columns are task, asignee, and due date.
task = ' '.join([line.strip() for line in action_items.strip().splitlines()]).split("|")
df = pd.DataFrame({'Task':task})
df[['Task','Assignee']] = df['Task'].str.split(':',expand=True)
df['Due Date'] = next_week.strftime('%m/%d/%y')

# Create email, enter email addresses, subject, and formatted email body.
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'esmith@example.com; mtrent@example.com; vpoonai@example.com;'
mail.Subject = 'Weekly Team Meeting Action Items - {d}'.format(d=d)
mail.HtmlBody = actions_items_email_template.format(d=d,df=df.to_html(index=False))
mail.Display(False)
