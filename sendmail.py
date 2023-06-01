import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os


def send_mail(receiver, file):
    body = '''Here is your certificate'''
    sender = 'envision.sitmng@gmail.com'
    password = 'gspzbhhdmhppizjo'
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = receiver
    message['Subject'] = 'Congratulations!!'

    message.attach(MIMEText(body, 'plain'))

    attachment = open(file, "rb")

    payload = MIMEBase("application", "octet-stream")
    payload.set_payload(attachment.read())

    # enconding the binary into base64
    encoders.encode_base64(payload)

    # add header with pdf name
    payload.add_header('Content-Disposition', 'attachment', filename='certificate.pdf')
    message.attach(payload)
    try:
        server = smtplib.SMTP('smtp.gmail.com: 587')
        server.starttls()
        server.login(sender, password)
        server.sendmail(message['From'], message['To'], message.as_string())
        server.quit()
        print(f'Mail Sent to {receiver}')
    except Exception as e:
        print(f"message not sent {receiver} because {e}")
