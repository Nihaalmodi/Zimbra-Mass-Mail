import os
import json
import base64

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from oauth2client.service_account import ServiceAccountCredentials

import mimetypes
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.multipart import MIMEMultipart
from email import encoders
from smtplib import SMTP

def create_message_with_attachment(sender, to, subject, message_text, files):
    """Create a message for an email.

    Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.

    Returns:
    An object containing a base64url encoded email object.
    """
    
    message = MIMEMultipart()
    message['Subject'] = subject
    message['From'] = sender
    message['To'] = ','.join(to)

    msg = MIMEText(message_text)
    message.attach(msg)
    if files is None:
        files = []
    for file in files:
        print(file)
        content_type, encoding = mimetypes.guess_type(file)
        print(content_type)
        if content_type is None or encoding is not None:
            content_type = 'application/octet-stream'
        main_type, sub_type = content_type.split('/', 1)
        # if main_type == 'text':
        #     fp = open(file, 'rb')
        #     msg = MIMEText(fp.read(), _subtype=sub_type)
        #     fp.close()
        if main_type == 'image':
            fp = open(file, 'rb')
            msg = MIMEImage(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'audio':
            fp = open(file, 'rb')
            msg = MIMEAudio(fp.read(), _subtype=sub_type)
            fp.close()
        else:
            fp = open(file, 'rb')
            msg = MIMEBase(main_type, sub_type)
            msg.set_payload(fp.read())
            encoders.encode_base64(msg)
            fp.close()
        filename = os.path.basename(file)
        msg.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(msg)

    return message.as_string()

def send_mail(sender, password, to, subject, message_text, file=None, SMTPserver='iitkgpmail.iitkgp.ac.in'):
    #TODO : Test Gmail
    if isinstance(to,str):
        to = [to]
    if isinstance(file,str):
        file = file.split(",")
    conn = None
    if (SMTPserver=='IITKGPMAIL (Zimbra)'):
        SMTPserver ='iitkgpmail.iitkgp.ac.in'
        conn = SMTP(SMTPserver)
    # elif (SMTPserver=='GMAIL'):
    #     conn = SMTP('smtp.gmail.com:587')
    #     conn.ehlo()
    #     conn.starttls()
    conn.set_debuglevel(1)
    conn.login(sender, password)
    # TODO: Error message for file not found
    message = create_message_with_attachment(sender, to, subject, message_text, file)
    try:
        conn.sendmail(sender, to, message)
        error_code = 200
    except Exception:
        error_code = 404
    finally:
        conn.quit()
    return error_code
