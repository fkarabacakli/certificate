# importing libraries
import os
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
import pandas as pd
from mailinfo import EMAIL_ADDRESS, EMAIL_PASSWORD

i = 1
df = pd.read_excel('data.xlsx')

for filename in os.listdir('certificates'):
    try:
        with open(os.path.join('certificates', filename), 'rb') as fp:
            file_data = fp.read()
            msg = EmailMessage()
            msg['Subject'] = "OBMYT Katılım Sertifikası"
            msg['From'] = formataddr(
                ('Okan Üniversitesi Bilgisayar Ve Yazılım Mühendisliği Kulübü', 'obmyt@gmail.com'))
            mail_to = filename.rsplit(".", 1)[0]
            msg['To'] = mail_to
            msg.set_content('Bu sertifika, Trendyol Tech Lead Emre Savcı''nın katılımıyla gerçekleşen "İş Hayatında Kariyer Yolculuğu" seminerine katılım gösterilerek almaya hak kazanılmıştır.')
            msg.add_attachment(file_data, maintype='application',
                               subtype='pdf', filename='KatılımSertifikası.pdf')
            print('File Name -> '+str(i)+" "+str(filename), end="")
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                smtp.send_message(msg)
            print(" Sending Success")
            df.loc[df['Mail']==mail_to, 'Sended'] = True
            i += 1
    except:
        print(str(mail_to) + " Sending Failed")

df.to_excel('data.xlsx', index=False)
