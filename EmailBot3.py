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
#this filters my messages to look for specific emails
yestMessages = inbox.Items.Restrict("[ReceivedTime] >= '" + yestFormat + "'")
yestMessages = yestMessages.Restrict("[Subject] = 'EOD Tracking'")

print(yestFormat)
print(yestMessages)



def messageHandler(dateFormat, subject):
    todayMessages = inbox.Items.Restrict("[ReceivedTime] >= '" + dateFormat + "'")
    todayMessages = todayMessages.Restrict("[Subject] = '" + subject + "'")
