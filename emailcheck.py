import time
import imaplib
import email
import sys
import re
from openpyxl import Workbook
from openpyxl import load_workbook
from config import *

# LOGGING-----------------------------------
te = open('Email_Log_' + time.strftime("%Y_%m_%d_%Hh_%Mm") + '.txt',
          'w')  # File where you need to keep the logs


class Unbuffered:
    def __init__(self, stream):
        self.stream = stream

    def write(self, data):
        self.stream.write(data)
        self.stream.flush()
        te.write(data)  # Write the data of stdout here to a text file as well


sys.stdout = Unbuffered(sys.stdout)
# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------

def display_visible_html_using_re(text):
    return(re.sub("(\<.*?\>)", "",text))

def read_email_from_gmail():
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        type, data = mail.search(None, 'ALL')
        mail_ids = data[0]

        id_list = mail_ids.split()
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        for i in range(latest_email_id,first_email_id, -1):
            typ, data = mail.fetch(i, '(RFC822)' )

            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    email_subject = msg['subject']
                    email_from = msg['from']
                    email_date= msg['date']
                    email_date=email_date[5:len(email_date)-6]
                    email_body='N/A'
                    if email_subject.startswith(Subject_Startswith):
                        wb = Workbook()
                        ws = wb.active
                        print 'From : ' + email_from + '\n'
                        print 'Subject : ' + email_subject + '\n'
                        print 'Date : ' + email_date + '\n'
                        try:
                            for part in msg.walk():
                                email_body = display_visible_html_using_re(part.get_payload())
                                print email_body
                        except:
                            print 'Email body couldnt be parsed'
                        ws.append([email_date, email_from, email_subject, email_body])
                        wb.save('Email_' + time.strftime("%Y_%m_%d_%Hh_%Mm") + '.xlsx')

    except Exception, e:
        print str(e)

read_email_from_gmail()