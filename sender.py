# Python code to send email to a list of
# emails from a spreadsheet
# https://myaccount.google.com/lesssecureapps disable 2FA and enable less secure apps
# import the required libraries
from time import sleep
import pandas as pd

import email
import smtplib
import ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# change these as per use
subject = "Summer internship"
sender_email = input("Type your email and press enter:")
password = input("Type your password and press enter:")

# establishing connection with gmail
context = ssl.create_default_context()
try:
    server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
    # server.ehlo()
    server.login(sender_email, password)

    # reading the spreadsheet
    email_list = pd.read_excel(
        'C:/Users/ousem/Documents/mail sender script/mails.xlsx')

    # getting the names and the emails
    names = email_list['NAME']
    emails = email_list['EMAIL']

    # iterate through the records
    for i in range(len(emails)):

        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = emails[i]
        message["Subject"] = subject

        # Add body to email
        body = f"Hello {names[i]}, CV mte3i mawjoud fil attachements?"
        message.attach(MIMEText(body, "plain"))
        filename = "Oussema TRABELSI CV.pdf"  # In same directory as script

        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        server.sendmail(sender_email, emails[i], text)
        # sleep for one minutes to avoid spam
        sleep(60)
except Exception as e:
    print(e)
finally:
    server.close()
