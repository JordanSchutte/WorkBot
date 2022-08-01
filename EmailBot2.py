import win32com.client
import re
import os
import time
import datetime
import tempfile
import pandas
import numpy
import string
from tempfile import mkstemp
import os.path as path
from datetime import date
from datetime import timedelta

today = date.today()
todayFormat = today.strftime('%m/%d/%Y')
path = r'Filepath'
#this stuff will change once testing is done
scriptPath, _ = os.path.split(__file__)
scratchPath = os.path.join(scriptPath, 'Scratch')
testfilePath = os.path.join(scratchPath, 'Testfile.xlsx')
temp = tempfile.TemporaryFile()
tempdir = tempfile.TemporaryDirectory(dir=scratchPath)
df = pandas.read_excel(testfilePath)
print(df)
#this creates a list of how many 'Sales Order' we need to append to data
rows, cols = df.shape
soList =[]
for i in range(0, rows):
    soList.append('Sales Order')
#this filters all the POs that require the manual look up in TC
#need a handler that can either save this txt and/or pass it to inbox bot
poList =[]
df['PO Number'] = df['PO Number'].map(str)
for po in df['PO Number']:
    if len(po) == 8:
        poList.append(po)
#this adds the list I just created to a new column on the left
salesOrderAdd = df.insert(0, 'Sales Order', value=soList)
#this turns the 'Sold To' column to the respective values needed for DB sync
df['Sold To'] = df['Sold To'].replace(numpy.nan, '0')
df['Sold To'] = df['Sold To'].replace('For Resale', '1')
#this turns the Zip codes into a string so we can manipulate the values
df['Ship To Postal Code'] = df['Ship To Postal Code'].map(str)
#this adds a leading 0 if the length is less than 5 and trims 4 digits if it = 9
df['Ship To Postal Code'] = df['Ship To Postal Code'].map(lambda x: x if len(x) > 4 else '0' + x)
df['Ship To Postal Code'] = df['Ship To Postal Code'].map(lambda x: x[:-4] if len(x) == 9 else x)

print(df)
print(poList)
#TODO phase3
#either copy and paste this data to clip board or overwrite tempfile
#find a dict that works with Access
#create connection with Access/Host PC this script might need to be on XPC
#open 'X QB Sync' table
#paste data from edited file into table
#ASSERT that the table accepted todays data
#run 'WM QB Sync' MACRO
#run ASSERT loop until 'WM QB Sync Lead Time Fetch v2' table is displayed
#parse data when sync is completed
