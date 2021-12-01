import win32com.client
import re
import os
import time
import datetime
import tempfile
from tempfile import mkstemp
import os.path as path
from datetime import date
from datetime import timedelta

today = date.today()
todayFormat = today.strftime('%m/%d/%Y')
yest = date.today() - datetime.timedelta(days = 1)
yestFormat = yest.strftime('%m/%d/%Y')
path = r'C:\Users\jordan\lpthw\Export'
#script path so this bot can run no matter where the host file exists
scriptPath, _ = os.path.split(__file__)
scratchPath = os.path.join(scriptPath, 'Scratch')
#temp file so I dont accumulate tons of files since we only need this data once
temp = tempfile.TemporaryFile()
tempdir = tempfile.TemporaryDirectory(dir=scratchPath)
print(tempdir)
#establish connection with Outlook
outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
inbox = outlook.GetDefaultFolder(6)

findMessage = getTodayMessages()
#this filters my messages to look for specific emails
def messageHandler(dateFormat, subject):
    todayMessages = inbox.Items.Restrict("[ReceivedTime] >= '" + dateFormat + "'")
    todayMessages = todayMessages.Restrict("[Subject] = '" + subject + "'")
    return todayMessages
#this locates my specified email, downloads the attachment as a tempfile
def attachmentHandler():
    try:
        for message in todayMessages:
            try:
                s = message.sender
                for attachment in message.Attachments:
                    attachment.SaveASFile(os.path.join(scratchPath, attachment.FileName))
                    print(f"attachment {attachment.FileName} from {s} saved")
            except Exception as e:
                print('Error when saving the attachment:' + str(e))
    except Exception as e:
        print('Error when processing email messages:' + str(e))
#I think this was a test
for root, dir, files in os.walk(scratchPath):
    for file in files:
        print(file)

def main():
    print('What email are we looking for?')
    print('Type 1 for EOD Tracking')
    print('Type 2 for TC SO Export')
    choice = input()
    if choice == '1':
        subject = 'EOD Tracking'
        messageHandler(yestFormat, subject)
        attachmentHandler()
    elif choice == '2':
        subject = 'TC SO Export (14102)'
        messageHandler(todayFormat, subject)
        attachmentHandler()

#this cleans up the tempfile to avoid clutter
tempdir.cleanup()
if __name__=="__main__":
    main()
