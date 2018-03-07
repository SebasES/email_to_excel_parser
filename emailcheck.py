import time
import imaplib
import email
import re
from openpyxl import Workbook
import textlog
from config import *

textlog.start_logging('Email_Log_' + time.strftime("%Y_%m_%d_%Hh_%Mm") + '.txt')
data=[]

def display_visible_html_using_re(text):
    return(re.sub("(\<.*?\>)", "",text))
# -------------------------------------------------
#
# Login
#
# ------------------------------------------------
try:
    mail = imaplib.IMAP4_SSL(SMTP_SERVER)
    mail.login(FROM_EMAIL,FROM_PWD)
    mail.select('inbox')
    type, data = mail.search(None, 'ALL')
except Exception, e:
    print 'Error retrieving E-Mails (Are you Online?): ' + str(e)
# -------------------------------------------------
#
# Parsing
#
# ------------------------------------------------
if len(data)>0:
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
                        print 'Email body couldnt be parsed... This doesnt seem to be a plain-text email'
                    ws.append([email_date, email_from, email_subject, email_body])
