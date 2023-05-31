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
    # put the email of the receiver here
    # receiver = 'adarshsavaligi@gmail.com'

    # Set up the MIME
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = receiver
    message['Subject'] = 'Congratulations!!'

    message.attach(MIMEText(body, 'plain'))

    binary_pdf = open(file, 'rb')

    payload = MIMEBase('application', 'octate-stream', Name=file)
    # payload = MIMEBase('application', 'pdf', Name=pdfname)
    payload.set_payload(binary_pdf.read())

    # enconding the binary into base64
    encoders.encode_base64(payload)

    # add header with pdf name
    payload.add_header('Content-Decomposition', 'attachment', filename=file)
    message.attach(payload)

    # use gmail with port
    session = smtplib.SMTP('smtp.gmail.com', 587)
    try:
    # enable security
        session.starttls()

        # login with mail_id and password
        session.login(sender, password)

        text = message.as_string()
        session.sendmail(sender, receiver, text)
        session.quit()

        print(f'Mail Sent to {receiver}')
    except Exception as e:
        print(f"message not sent {receiver} beacuse {e}")

