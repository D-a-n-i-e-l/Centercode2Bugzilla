# I have created a slack app that will, in the final form of this project, tag a TSE and send the bug url to a slack channel with some added information.

import requests

token = "//"
channel_id = "//"
post_url = 'https://slack.com/api/chat.postMessage'

r = requests.post(post_url, data={"token": token, "channel": channel_id,
"text": "------------------------------\n" 
+ ID_finish + summ + "Bug URL: " + url + "\n" + "<@"+ slack_name +">" 
+ "\n------------------------------"})



# Ultimately will send a slack message that looks like this:

#------------------------------
#DIS00089

# This is the title of the bug parsed from the issue.

#Bug URL: http://tinyurl.com/compressedURL
#@TSE_User(parsed from template)
#------------------------------
