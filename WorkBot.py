import pyautogui as pog
import time
import os
import pytesseract
import argparse
import win32com.client as win32
import pandas as pd
import csv
import sys
import os.path as path
from datetime import date
from datetime import timedelta, date

path = r'C:\Users\Jordan\Python\Resources'
today = date.today()
delivery = today+timedelta(days=10)
ship = today+timedelta(days=7)
absShip = today-timedelta(days=1)
absDelivery = today+timedelta(days=3)
todayFormat = today.strftime('%m/%d/%Y')
deliveryFormat = delivery.strftime('%m/%d/%Y')
shipFormat = ship.strftime('%m/%d/%Y')
absShipFormat = absShip.strftime('%m/%d/%Y')
absDeliveryFormat = absDelivery.strftime('%m/%d/%Y')
exportFormat = today.strftime('%m-%d-%Y')

def buildFileMap(path):
    fileMap = {}
    root = None
    absPath = None
    for aPath, dirs, files in os.walk(path):
        for file in files:
            root, ext = os.path.splitext(file)
            absPath = os.path.join(aPath, file)
            os.path.normpath(absPath)
            fileMap[root] = absPath
    return fileMap

fileMap = buildFileMap(path)

def reLogin():
    pog.hotkey('ctrl', 'w')
    pog.hotkey('ctrl', 't')
    leftClickLabel('TCWeb')
    time.sleep(3)
    leftClickLabel('LoginBar')
    leftClickLabel('EmailAdd')
    leftClickLabel('NextEvent')
    time.sleep(1)
    leftClickLabel('LoginEvent')

def findLabel(labelName):
    fname = fileMap[labelName]
    assert(fname)
    try:
        label = pog.locateOnScreen(fname, confidence = 0.9)
    except pog.ImageNotFoundException:
        print('Next task')
    return label

def leftClickLabel(labelName):
    print('Left clicking label:', labelName)
    fname = pog.center(findLabel(labelName))
    fnamex, fnamey = fname
    pog.click(fnamex, fnamey)

def ediCheck(edi):
    if edi == '1':
        pog.write(' (855)')
    elif edi == '2':
        pog.write(' (856)')
    else:
        sys.exit('EDI is not valid.')

def poPhase1():
    leftClickLabel('Red')
    pog.press('enter')
    pog.write('Item Accepeted')
    pog.press(['enter','tab','tab'])
    pog.write('Each')
    pog.press('enter')
    time.sleep(1)
    pog.scroll(-50)
    time.sleep(1)

def handlerInbox(label, edi):
    assert(findLabel('InboxDrop'))
    labelDocNum = findLabel('DocNum')
    POx = labelDocNum.left + labelDocNum.width
    _ , POy = pog.center(label)
    pog.moveTo(POx, POy)
    pog.drag(-100, 0, button='left')
    pog.hotkey('ctrl', 'c')
    leftClickLabel('Turn')
    time.sleep(1)
    leftClickLabel('Async')
    pog.hotkey('ctrl', 'v')
    ediCheck(edi)
    leftClickLabel('StartEvent')
    time.sleep(2)
    pog.scroll(-500)
    time.sleep(1)
    if edi == '1':
        leftClickLabel('POAck')
    if edi == '2':
        leftClickLabel('ShipNotice')
    leftClickLabel('GreenBox')
    time.sleep(1)
    leftClickLabel('OKEvent')
    time.sleep(2)
    leftClickLabel('Refresh')
    time.sleep(3)

def handlerOutbox(label):
    assert(findLabel('OutboxDrop'))
    labelDocNum = findLabel('DocNum')
    POx = labelDocNum.left + labelDocNum.width
    _ , POy = pog.center(label)
    pog.moveTo(POx, POy)
    pog.drag(-100, 0, button='left')
    pog.hotkey('ctrl', 'c')
    pog.doubleClick()
    time.sleep(5)
    leftClickLabel('RedBar')
    pog.write('Accepted')
    pog.press('enter')
    pog.scroll(-100)
    time.sleep(1)
    leftClickLabel('RedBar')
    time.sleep(1)
    pog.write(todayFormat)
    leftClickLabel('RedItems')
    time.sleep(3)
    greenEach = findLabel('HighlightedE')
    pog.moveTo(greenEach)
    pog.keyDown('shift')
    pog.scroll(-7000)
    pog.keyUp('shift')
    time.sleep(2)
    label = findLabel('Red')
    while label:
        poPhase1()
        label = findLabel('Red')
    pog.scroll(1000)
    pog.keyDown('shift')
    pog.scroll(-3000)
    pog.keyUp('shift')
    time.sleep(2)
    label = findLabel('Yellow')
    while label:
        leftClickLabel('Yellow')
        pog.write(deliveryFormat)
        pog.press('enter')
        time.sleep(1)
        pog.scroll(-50)
        time.sleep(1)
        label = findLabel('Yellow')
    pog.scroll(1000)
    pog.keyDown('shift')
    pog.scroll(-3000)
    pog.keyUp('shift')
    time.sleep(2)
    label = findLabel('MiniRed')
    while label:
        leftClickLabel('MiniRed')
        pog.press('Enter')
        pog.write('English')
        pog.press('Enter')
        time.sleep(1)
        pog.scroll(-50)
        time.sleep(1)
        label = findLabel('MiniRed')
    pog.scroll(1000)
    leftClickLabel('RedShipping')
    time.sleep(1)
    leftClickLabel('YellowBar')
    pog.write(deliveryFormat)
    pog.press('Tab')
    pog.write(shipFormat)
    leftClickLabel('RedBar')
    pog.write('Per Contract')
    leftClickLabel('SaveNClose')
    time.sleep(3)
    leftClickLabel('Send')
    time.sleep(2)
    leftClickLabel('AsyncSend')
    pog.hotkey('ctrl', 'v')
    ediCheck(edi)
    leftClickLabel('StartEvent')
    time.sleep(2)

def searchPO(label):
    leftClickLabel('SearchIcon')
    time.sleep(3)
    leftClickLabel('SearchBar')
    pog.hotkey('ctrl', 'v')
    leftClickLabel('SearchFunction')
    time.sleep(2)
    leftClickLabel('Received')
    leftClickLabel('Options')
    time.sleep(1)
    leftClickLabel('MoveToInbox')
    time.sleep(1)
    leftClickLabel('OK16425')
    leftClickLabel('OKEvent')
    time.sleep(3)
    leftClickLabel('Outbox')
    time.sleep(2)

def handlerShip(label):
    assert(findLabel('OutboxDrop'))
    labelAltDoc = findLabel('AltDoc')
    POx = labelAltDoc.left + labelAltDoc.width
    _ , POy = pog.center(label)
    pog.moveTo(POx, POy)
    pog.drag(-100, 0, button='left')
    pog.hotkey('ctrl', 'c')
    pog.doubleClick()
    time.sleep(5)
    leftClickLabel('Dates')
    time.sleep(1)
    leftClickLabel('RedBar')
    pog.write(absShipFormat)
    leftClickLabel('YellowBar')
    pog.write(absDeliveryFormat)
    leftClickLabel('FOB')
    time.sleep(1)
    leftClickLabel('RedBar')
    pog.write('Per Contract')
    leftClickLabel('Addresses')
    time.sleep(2)
    leftClickLabel('ShipFrom')
    time.sleep(1)
    leftClickLabel('RedBar')
    pog.write('BUISNESS')
    time.sleep(1)
    pog.press('Tab')
    pog.write('ADDRESS')
    time.sleep(1)
    pog.press(['Tab','Tab'])
    pog.write('CITY')
    time.sleep(1)
    pog.press('Tab')
    pog.write('STATE')
    time.sleep(1)
    pog.press('Tab')
    pog.write('ZIPCODE')
    time.sleep(1)
    pog.press(['Tab','Tab'])
    pog.write('SCAC')
    time.sleep(1)
    pog.press('Tab')
    pog.write('FXFE')
    leftClickLabel('SaveNClose')
    time.sleep(3)
    leftClickLabel('Send')
    time.sleep(2)
    leftClickLabel('AsyncSend')
    pog.hotkey('ctrl', 'v')
    ediCheck(edi)
    leftClickLabel('StartEvent')
    time.sleep(3)
    leftClickLabel('Refresh')
    time.sleep(3)

print('What must be done for these POs?')
print('Type 1 for 855, or 2 for 856')
edi = input()

#TODO: current loops always thinks there is one more than there actually is.???

def main(): #this is new loop method, make sure it works
    itemsCompletedInbox = 0
    itemsCompletedOutbox = 0
    label = findLabel('TargetTag')
    while label:
        handlerInbox(label, edi)
        itemsCompletedInbox += 1
        label = findLabel('TargetTag')
    leftClickLabel('Outbox')
    time.sleep(3)
    if edi == '1':
        label = findLabel('TargetTag')
        while label:
            handlerOutbox(label)
            searchPO(label)
            itemsCompletedOutbox += 1
            label = findLabel('TargetTag')
    elif edi == '2': #test this next
        label = findLabel('TargetTag')
        while label:
            handlerShip(label)
            itemsCompletedOutbox += 1
            label = findLabel('TargetTag')
    print('Items moved from Inbox:', itemsCompletedInbox)
    print('Items sent from Outbox', itemsCompletedOutbox)
    sys.exit('Work Complete')

if __name__ == '__main__':
    main()
