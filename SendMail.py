# -*- coding: utf-8 -*-
"""
Created on Wed Jul 14 22:35:30 2021

@author: bhavya boppana
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

df = pd.read_excel("C:\\Users\\bhavya boppana\\Downloads\\WORLD_CUP.xlsx")


fromaddr = "testmailtableau@gmail.com"
print(df)
toaddr = str(df.iloc[0,7])
   
# instance of MIMEMultipart
msg = MIMEMultipart()
  
# storing the senders email address  
msg['From'] = fromaddr
  
# storing the receivers email address 
msg['To'] = str(toaddr)
  
# storing the subject 
msg['Subject'] = "Tests on the tableau workbook have been done"
  
# string to store the body of the mail
body = "Checkout the results at the URL '\\20.186.37.98\TableauTestResults\TestResults'"
  
# attach the body with the msg instance
msg.attach(MIMEText(body, 'plain'))
  
# open the file to be sent 
filename = "WORLD_CUP.xlsx"
attachment = open("C:\\Users\\bhavya boppana\\Downloads\\WORLD_CUP.xlsx", "rb")
  
# instance of MIMEBase and named as p
p = MIMEBase('application', 'octet-stream')
  
# To change the payload into encoded form
p.set_payload((attachment).read())
  
# encode into base64
encoders.encode_base64(p)
   
p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
  
# attach the instance 'p' to instance 'msg'
msg.attach(p)
  
# creates SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)
  
# start TLS for security
s.starttls()
  
# Authentication
s.login(fromaddr, "Admin@123")
  
# Converts the Multipart msg into a string
text = msg.as_string()
  
# sending the mail
s.sendmail(fromaddr, toaddr, text)
  
# terminating the session
s.quit()