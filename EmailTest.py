#My code

outlook = win32.Dispatch('outlook.application').GetNamespace('MAPI')
outlookInbox = mapi.GetDefaultFolder(6)
exportMsg = msg.Restrict("[Subject] = 'Foo'")
outputDir = r'Directorypath' #This needs to save on local, not host

def exportSave():
    for exportMsg in all_outlookInbox:
        print(msg.subject)
        try:
            for att in msg.Attachements:
                print(att.FileName)
                att.SaveASFile(os.path.join(outputDir + '\\' + str(att))) #this needs to save to local
                print(f'attachment {att.FileName} from message has been saved.')
        except Exception as e:
            print('error when saving the attachment:' + str(e))

def walmartExport():
    todayWorkbook = openpyx1.load_workbook(filename= 'Foo ' + exportFormat + ' Bar - Baz.xlsx')
    activeSheet = todayWorkbook.active
    todayWorkBook.insert_cols(1)


#Class concept

import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from mimetypes import guess_type
from email.encoders import encode_base64
from getpass import getpass
from smtplib import SMTP


class Email(object):
    def __init__(self, from_, to, subject, message, message_type='plain',
                 attachments=None):
        self.email = MIMEMultipart()
        self.email['From'] = from_
        self.email['To'] = to
        self.email['Subject'] = subject
        text = MIMEText(message, message_type)
        self.email.attach(text)
        if attachments is not None:
            for filename in attachments:
                mimetype, encoding = guess_type(filename)
                mimetype = mimetype.split('/', 1)
                fp = open(filename, 'rb')
                attachment = MIMEBase(mimetype[0], mimetype[1])
                attachment.set_payload(fp.read())
                fp.close()
                encode_base64(attachment)
                attachment.add_header('Content-Disposition', 'attachment',
                                      filename=os.path.basename(filename))
                self.email.attach(attachment)
