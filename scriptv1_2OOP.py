import win32com.client
import urllib.parse
from pyshorteners import Shortener
import requests

class CC_to_BZscript:

	def get_body(self):
		application = win32com.client.Dispatch('Outlook.Application')
		namespace = application.GetNamespace('MAPI')
		inbox_folder = namespace.GetDefaultFolder(6) 
		inbox = inbox_folder.Folders
		example_folder = inbox["Centercode Projects"]
		example_example_folder = example_folder.Folders["Folder Name"]
		messages = example_example_folder.Items
		messages.Sort("[ReceivedTime]", False)
		message = messages.GetLast()
		query = message.Body
		return query


	def dothesplits(self):
		query = CC_to_BZscript.get_body(self)
		query = query.split("Triage")[0]
		summary = query.split("Summary:")[1]
		summary_mid = summary.split("Details and")[0] 
		summary_fin = urllib.parse.quote(summary_mid)
		return summary_fin
		
	def return_body(self):
		query = CC_to_BZscript.get_body(self)
		formatted = urllib.parse.quote(query)
		return formatted

	def construct_url(self):
		url = "Bugzilla Template with concactenated variables"
		for i in range(0,100):
			while True:		
				try:
					shorten = Shortener('Tinyurl')
					url = shorten.short(url)
				except requests.exceptions.ReadTimeout:
					continue
				break
		return url



	def slackbot(self):
		url = CC_to_BZscript.construct_url(self)
		query = CC_to_BZscript.get_body(self)
		name_beg = query.split("Triage: ")[1]
		name = name_beg.split("\r")[0]
		slack_name = name[0] + name.split(" ")[1]
		slack_name = slack_name.lower()
		summary_non_utf = query.split("Summary:")[1]
		summ = summary_non_utf.split("Details and")[0] 
		ID_start = query.split("ID :")[1]
		ID_finish = ID_start.split("Incident")[0]


		token = "???"


		channel_id = "???"



		r = requests.post('https://slack.com/api/chat.postMessage', 
			data={
			"token": token,
			"channel": channel_id,
			"text": 
				"------------------------------\n" 
      				+ ID_finish + summ + "Bug URL: " + url + "\n" + "<@"+ slack_name +">" + 
      				"\n------------------------------"
				}
			)

 
