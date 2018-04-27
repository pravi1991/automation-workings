import imaplib_connect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import configparser
import os
import datetime
import time

today = str(datetime.datetime.now().date())
setting_path = os.path.realpath('.settings.txt')
config = configparser.ConfigParser()
config.read(setting_path)
username = config.get('account_juno', 'username')
password = config.get('account_juno', 'password')
toaddr = ["praveen.tt@junotele.com" ]
Subject = 'JSReport for Dynamic YES/NO'
Message = "Hi,\n\nPlease find the enclosed JSecure Dynamic YES/NO report. \n\n(Dynamic Yes/No Team)"  # for " + str(datetime.datetime.now().hour) + " th hour"
filename = 'Voda_Report_' + today + '.xlsx'
WAIT = 60


def send_report():
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = ", ".join(toaddr)
    msg['Subject'] = Subject
    # msg['CC'] = 'rajeev.m@junotele.com'

    msg.attach(MIMEText(Message, 'plain'))
    try:
        print("\t" + filename)
        attachment = open(os.path.realpath(filename), "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)
        print("\tFile Attached")

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)  # if tls is not suppor use SMTP_SSL with port 465.
        # server.starttls()
        server.ehlo()
        server.login(username, password)
        text = msg.as_string()
        print("\tSending to\t::", toaddr)
        server.sendmail(username, toaddr, text)
        server.quit()
        print("\tSending Done")

    except FileNotFoundError:
        print("\tFile attach failed.. Check if file is present")

def send_ack(Subject,Message):
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = ", ".join(['praveen.tt@junotele.com','rajeev.m@junotele.com'])
    msg['Subject'] = Subject
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)  # if tls is not suppor use SMTP_SSL with port 465.
    server.ehlo()
    server.login(username, password)
    server.sendmail(username, ['praveen.tt@junotele.com','rajeev.m@junotele.com'], text)
    server.quit()


with imaplib_connect.open_connection() as mail:
    mail.select(mailbox='COMMANDS', readonly=False)
    typ, msg_ids = mail.search(None, '(SUBJECT "{}" UNSEEN)'.format(
        'Stop emails to John'))
    msg_id = msg_ids[0].decode('ascii').split(' ')
    '''-------------Never Delete these 'print' statements--------------'''
    if msg_id != ['']:
        stop_id = msg_id[0]
        print('STOP COMMAND RECEIVED\nSTOP ID \t::', stop_id)
        print('All external mails STOPPED')
        for i in range(WAIT):
            print('Waiting for START command:', i, ' second(s)', end='\r')
            time.sleep(1)

        mail.select(mailbox='COMMANDS', readonly=False)
        typ, msg_ids = mail.search(None, '(SUBJECT "{}" FROM "nikhil.gupta@junotele.com" UNSEEN)'.format(
            'Start emails to John'))
        msg_id = msg_ids[0].decode('ascii').split(' ')
        if msg_id != ['']:
            print("\r\nRESEND COMMMAND RECEIVED\nSEND ID \t::", msg_id[0])
            mail.store(stop_id, '+FLAGS', ('\SEEN'))
            mail.store(msg_id[0], '+FLAGS', ('\SEEN'))
            print('MARKED SEEN\t::', stop_id, msg_id[0])
            print("All external mails RESUMED")
            send_report()
        else:
            print('No resend command received in . Continuing STOP state..!!\n')

    else:
        send_report()
