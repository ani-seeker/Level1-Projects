import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path


def send_email(email_recipient,
               email_subject,
               email_message,
               attachment_location1 = '',
               attachment_location2 = ''):

    email_sender = 'anirudh.gupta@blackwoods.com.au'

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))

    if attachment_location1 != '':
        filename = os.path.basename(attachment_location1)
        attachment = open(attachment_location1, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename= %s" % filename)
        msg.attach(part)

    if attachment_location2 != '':
        filename = os.path.basename(attachment_location2)
        attachment = open(attachment_location2, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename= %s" % filename)
        msg.attach(part)


    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login('anirudh.gupta@blackwoods.com.au', 'Arrow_09876')
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('email sent from :', email_sender,'to :',  email_recipient)
        server.quit()
    except:
        print("SMPT server connection error")
    return True

send_email('anirudh.gupta@blackwoods.com.au',
           'Report - Daily Updates - STIBO vs IICE Comparison',
           'Please find attached the daily report',
           r'C:\Users\axg03\IICE_Part_Level_Comparison_with_STIBO_Daily.xlsx',
           r'C:\Users\axg03\IICE_Vendor_Part_Level_Comparison_with_STIBO_Daily.xlsx')
