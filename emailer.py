"""
Class to send email
Author: Ryan Morlando
Created:
Updated:
V1.0.0
Patch Notes:

To Do:

"""
import re
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


def send_email(credentials: dict, subject: str, receivers: list,  body: str,
               body_type: str = None, attachment_paths: list = None) -> bool:
    """

    :param credentials: dict {'host': 'email host', 'port': 'port number', 'login': 'email login', 'pwd': 'email pwd',
                                'sender': 'sender email'}
    :param subject: str
    :param receivers: list of receivers email
    :param body: str, can be html if body_type = 'html' is specified
    :param body_type: if html, pass 'html'
    :param attachment_paths: list of any attachments. Must be full path to file
    :return bool: True if the message was sent. False if there was an error
    """
    try:
        host = credentials['host']
        port = credentials['port']
        login = credentials['login']
        login_pwd = credentials['pwd']
        sender = credentials['sender']
    except KeyError:
        print('Credentials dict incomplete: Missing one of (host, port, login, pwd, or sender)')
        return False

    msg = MIMEMultipart()
    msg.add_header('From:', sender)
    msg.add_header('To:', ",".join(receivers))
    msg.add_header('Subject', f"{subject}")

    # add attachments
    if attachment_paths is not None:
        for file in attachment_paths:
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(file, "rb").read())
            encoders.encode_base64(part)
            file_name = re.sub(r'(C:.*\\)*(.*/)*', '', file)
            part.add_header('Content-Disposition', 'attachment; filename="' + file_name + '"')
            msg.attach(part)

    # add body
    if body_type is None:
        msg_body = MIMEText(body)
    else:
        msg_body = MIMEText(body, body_type)

    msg.attach(msg_body)

    mailserver = smtplib.SMTP_SSL(host, port)
    mailserver.login(login, login_pwd)

    try:
        mailserver.sendmail(sender, receivers, msg.as_string())
        mailserver.quit()
        return True
    except smtplib.SMTPException as e:
        print("Error: unable to send email")
        print(e)
        mailserver.quit()
        return False
