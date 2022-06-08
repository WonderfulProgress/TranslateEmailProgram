#!/usr/bin/env python
#pip install googletrans
#pip install googletrans==3.1.0a0
#pip install xlrd
#pip install pandas
#pip install openpyxl

import smtplib
from googletrans import Translator
from email.mime.text import MIMEText
from getpass import getpass
import pandas as pd
import xlrd
import openpyxl
import re

#Sender email information (Only support gmail)

print("Your username")
username = input("Username: ")

print("Your password")
password = getpass()

print("Your email address")
sender = input("Email: ")

smtp_ssl_host = 'smtp.gmail.com' 
smtp_ssl_port = 465

#Message (Txt file in same folder)
with open('email.txt', 'r') as file:
    data = file.read()

#Translator (Only support three language for now, more can be added)
#{'af': 'afrikaans', 'sq': 'albanian', 'am': 'amharic', 'ar': 'arabic', 'hy': 'armenian', 'az': 'azerbaijani', 'eu': 'basque', 'be': 'belarusian', 'bn': 'bengali', 'bs': 'bosnian', 'bg': 'bulgarian', 'ca': 'catalan', 'ceb': 'cebuano', 'ny': 'chichewa', 'zh-cn': 'chinese (simplified)', 'zh-tw': 'chinese (traditional)', 'co': 'corsican', 'hr': 'croatian', 'cs': 'czech', 'da': 'danish', 'nl': 'dutch', 'en': 'english', 'eo': 'esperanto', 'et': 'estonian', 'tl': 'filipino', 'fi': 'finnish', 'fr': 'french', 'fy': 'frisian', 'gl': 'galician', 'ka': 'georgian', 'de': 'german', 'el': 'greek', 'gu': 'gujarati', 'ht': 'haitian creole', 'ha': 'hausa', 'haw': 'hawaiian', 'iw': 'hebrew', 'hi': 'hindi', 'hmn': 'hmong', 'hu': 'hungarian', 'is': 'icelandic', 'ig': 'igbo', 'id': 'indonesian', 'ga': 'irish', 'it': 'italian', 'ja': 'japanese', 'jw': 'javanese', 'kn': 'kannada', 'kk': 'kazakh', 'km': 'khmer', 'ko': 'korean', 'ku': 'kurdish (kurmanji)', 'ky': 'kyrgyz', 'lo': 'lao', 'la': 'latin', 'lv': 'latvian', 'lt': 'lithuanian', 'lb': 'luxembourgish', 'mk': 'macedonian', 'mg': 'malagasy', 'ms': 'malay', 'ml': 'malayalam', 'mt': 'maltese', 'mi': 'maori', 'mr': 'marathi', 'mn': 'mongolian', 'my': 'myanmar (burmese)', 'ne': 'nepali', 'no': 'norwegian', 'ps': 'pashto', 'fa': 'persian', 'pl': 'polish', 'pt': 'portuguese', 'pa': 'punjabi', 'ro': 'romanian', 'ru': 'russian', 'sm': 'samoan', 'gd': 'scots gaelic', 'sr': 'serbian', 'st': 'sesotho', 'sn': 'shona', 'sd': 'sindhi', 'si': 'sinhala', 'sk': 'slovak', 'sl': 'slovenian', 'so': 'somali', 'es': 'spanish', 'su': 'sundanese', 'sw': 'swahili', 'sv': 'swedish', 'tg': 'tajik', 'ta': 'tamil', 'te': 'telugu', 'th': 'thai', 'tr': 'turkish', 'uk': 'ukrainian', 'ur': 'urdu', 'uz': 'uzbek', 'vi': 'vietnamese', 'cy': 'welsh', 'xh': 'xhosa', 'yi': 'yiddish', 'yo': 'yoruba', 'zu': 'zulu', 'fil': 'Filipino', 'he': 'Hebrew'}

translator = Translator()
cn_result = translator.translate(data, src='en', dest='zh-cn')
ja_result = translator.translate(data, src='en', dest='ja')
es_result = translator.translate(data, src='en', dest='es')

# read parents information from excel file
df = pd.read_excel('parents_info.xlsx')

server = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
server.login(username, password)

#Select group of parents the email is sending to
print("Email Group Type: 1. Volunteer  2. Covid Info  3. General")
ans = int(input("Number: "))

if ( ans == 1 ):
    filter = df[df['Volunteer'] == 'Yes']
elif ( ans == 2 ):
    filter = df[df['CovidInfo'] == 'Yes']
else:
    filter = df[df['Number'] > 0 ]

#input email subject
print("Email subject")
subject = input("Subject: ")


# read from parents list for sending list and send out email

email_regex = re.compile(r"[^@]+@[^@]+\.[^@]+")

for row in filter.values:
    email = f'{row[3]}'
    if not email_regex.match(email):
        continue
    language = f'{row[4]}'
    if ( language == "Chinese" ):
        final_message = data + cn_result.text
    elif ( language == "Japanese" ):
        final_message = data + ja_result.text
    elif ( language == "Spanish" ):
        final_message = data+ es_result.text
    else:
        final_message = data

    msg = MIMEText(final_message)
    msg['To'] = ', '.join(email)
    msg['Subject'] = subject
    msg['From'] = sender
    server.sendmail(sender, email, msg.as_string())

server.quit()

print("Emails were sent out successfully!")

