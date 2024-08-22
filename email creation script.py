import pandas as pd
import requests as req
import random
import string
import os
import smtplib
from email.message import EmailMessage
from datetime import datetime
now = datetime.now()
date_time = now.strftime("%m-%d-%Y-%H-%M-%S")
fnam = "Email_Record_" + date_time+".xlsx"


def gen_pass(num):
    # get random password pf length 8 with letters, digits, and symbols
    characters = string.ascii_letters + string.digits + string.punctuation
    password = ''.join(random.choice(characters) for i in range(num))
    return password


def create_email(email, pswd):
    url = "https://flourisense.in:2083/execute/Email/add_pop?email={email}&password={pswd}".format(
        email=email, pswd=pswd)

    headers = {
        'Authorization': 'cpanel sumeetbalwade:MUPSWXAT4TBBD0NPYQDO382EJHE9G1JR'}

    r = req.get(url, headers=headers)

    res = r.json()

    return res


def email_template(remail, rpass):

    file_path = "email_template.html"

    if os.path.isfile(file_path):
        text_file = open(file_path, "rt")

        data = text_file.read()

        text_file.close()

        data = data.replace("#email#", remail)
        data = data.replace("#password#", rpass)
        return data


def send_email(to, html):
    EMAIL_ADDRESS = "noreply@flourisense.in"
    EMAIL_PASSWORD = "Flourisense@123"

    msg = EmailMessage()
    msg['Subject'] = 'FLOURISENSE SERVER EMAIL CREDENTAILS'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to
    msg['Cc'] = ["sumeet.balwade@flourisense.in",
                 "deepchakkar.flourisense@gmail.com", "hr.flourisense@gmail.com"]

    msg.set_content('This is a plain text email')

    msg.add_alternative(html, subtype='html')

    with smtplib.SMTP_SSL('mail.flourisense.in', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)


# main process
pd.options.mode.chained_assignment = None
set = pd.read_excel('Email1.xlsx', header=0)
cols = set.columns

for i in range(set.shape[0]):

    name = set["Name"][i].lower()
    name2 = name.split()[0]+"." + name.split()[-1]
    email = name2+"@flourisense.in"
    password = gen_pass(9)
    pmail = set["Email"][i].lower()

    try:
        res = create_email(email, password)
        if(res["status"] == 1):
            send_email(pmail, email_template(email, password))
            set["Official Email"][i] = email
            set["CpanelPassword"][i] = password
            set["Remarks"][i] = "Created"
            print(email + " Created")

        else:
            set["Remarks"][i] = "Not Created"
            set["Data"][i] = ' '.join([str(elem) for elem in res["errors"]])
            print(email + " Not Created")
    except:
        print("An exception occurred")


set.to_excel(fnam)
