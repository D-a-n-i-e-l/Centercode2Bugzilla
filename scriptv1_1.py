import win32com.client
import urllib.parse
import webbrowser
from pyshorteners import Shortener
import requests

# Initializes win32 client to interact with Microsoft Outlook
application = win32com.client.Dispatch('Outlook.Application')
namespace = application.GetNamespace('MAPI')

# Creats a MAPI folder object of the main Inbox folder. 6 is the number for the main inbox
inbox_folder = namespace.GetDefaultFolder(6) 

# Had to create multiple objects of subfolders to get to specific directory
inbox = inbox_folder.Folders
sub_folder = inbox["Projects"]
sub_sub_folder = sub_folder.Folders["Example Project"]


# Used the Items method to parse individual email files within the folder
messages = sub_sub_folder.Items

# The Sort function will sort your messages by their ReceivedTime property, from the most recently received to the oldest
# If you use False instead of True, it will sort in the opposite direction: ascending order, from the oldest to the most recent
messages.Sort("[ReceivedTime]", False)

# GetLast() will retrieve the last email recieved per your sorting
message = messages.GetLast()
# Body method creates an object that just has the body of the email
query = message.Body


#-------------------------------------------------------------------
# I parsed certain information out to insert to my Bug URL. I wanted to parse a summary statement and a body/details statement.

# Turn body+summary+details into UTF-8 string so Webbrowser and Bugzilla can interpret it
formatted = urllib.parse.quote(query)

# Take the summary of the issue report so we can put it into the summary of the bug. Use slices to acheive this. Then encode summary.
summary = query.split("Summary:")[1]
summary_mid = summary.split("Details and")[0] 
summary_final = urllib.parse.quote(summary_mid)

# Concactinate url with formatted strings

long_url = "First part of bug template URL" + formatted + "Other part of bug temp URL"+ summary_final + "Final part of URL"


#------------------------------------------------------------------------

# Shorten URL since Chrome cuts URL off around 2,000 characters(and these are some long urls!). Have to sign into bugzilla beforehand for this to work.
for i in range(0,100):
  while True:
    try:
      shorten = Shortener('Tinyurl')
      url = shorten.short(long_url)
    except requests.exceptions.ReadTimeout
      continue
    break

# --------------------------------------------


# Open webbroswer with url just created
webbrowser.open(url)
