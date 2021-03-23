import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email.mime.application
import pandas as pd
import time

class Empfänger:
    def __init__(self, Nachname, Vorname, Email, Anrede):
        self.nachname = Nachname
        self.vorname = Vorname
        self.email = Email
        self.anrede = Anrede

def GetReceivers():
    filepath = 'Empfänger'
    Receivers = []

    with open(filepath, 'r') as lines:
        for line in lines:
            line = line.strip('\n')
            result = line.split(',')
            Receivers.append(Empfänger(result[0].strip(), result[1].strip(), result[2].strip(), result[3].strip()))
    return Receivers

def Anrede(Message, Receiver):
    if Receiver.anrede == 'Herr':
        return Message.format('geehrter Herr',Receiver.nachname)
    if Receiver.anrede == 'Frau':
        return Message.format('geehrtee Frau',Receiver.nachname)

def GetMessage():
    return  open('Nachricht', 'r').read()

start = time.time()
#Initialize
sender = 'YOUR_MAIL'
password = 'PASSWORD'
cc = 'CC_MAIL'
subject = 'SUB'

receiver_list = GetReceivers()

#start server
server = smtplib.SMTP('smtp.office365.com', 587)
server.starttls()
server.login(sender, password)



Message = GetMessage()

for receiver in receiver_list:
    
    message = Anrede(Message, receiver)

    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = receiver.email
    msg['Subject'] = subject
    msg['CC'] = cc
    msg.attach(MIMEText(message, 'plain'))
    
    filename = 'Anhang.pdf' #path to file
    fo=open(filename,'rb')
    attach = email.mime.application.MIMEApplication(fo.read(),_subtype="pdf")
    fo.close()
    attach.add_header('Content-Disposition','attachment',filename=filename)
    msg.attach(attach)
    msg.attach(attach)

    text = msg.as_string()

    server.sendmail(sender, receiver.email, text)

server.quit()

fin = time.time()
print('Email is transmitted within %.3f'%(fin-start))

