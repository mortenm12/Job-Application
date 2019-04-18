import os
import sys 
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename
import csv
import datetime
import getpass

now = datetime.datetime.now()
#arguments: companyName, companyEmail, posistion, additionalDescription 

TITLELINE = 18
DESCLINE = 24
SERVER = "smtp-mail.outlook.com"
FROM = "mortenm12@hotmail.com"
PORT = 587

def eoaReplace(text):
    text = text.replace("æ", "\\ae{}")
    text = text.replace("ø", "\\o{}")
    text = text.replace("å", "\\aa{}")

    text = text.replace("Æ", "\\AE{}")
    text = text.replace("Ø", "\\O{}")
    text = text.replace("Å", "\\AA{}")
    return text

def replaceLine(file_name, line_num, text):
    lines = open(file_name, 'r').readlines()
    lines[line_num] = text
    out = open(file_name, 'w')
    out.writelines(lines)
    out.close()

def cleanUpLatex(name):
    name = name.replace(" ", "")
    os.remove("tex/" + name + ".aux")
    os.remove("tex/" + name + ".log")
    os.remove("tex/" + name + ".pdf")

    replaceLine("tex/main.tex", TITLELINE, eoaReplace("\section*{Ansøgning}\n"))
    replaceLine("tex/main.tex", DESCLINE, "\n")

def addToLatex(name, desc):
    replaceLine("tex/main.tex", TITLELINE, eoaReplace("\section*{Ansøgning - " + name +"}\n"))
    replaceLine("tex/main.tex", DESCLINE, eoaReplace(desc) +"\n")

def sendEmail(password, email, subject, text, f):
    f = f.replace(" ", "")
    server = smtplib.SMTP(SERVER, PORT)
    server.ehlo()
    server.starttls() 
    server.login(FROM, password)
    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg.attach(MIMEText(text))
    
    with open("tex/"+f, "rb") as fil: 
        ext = f.split('.')[-1:]
        attachedfile = MIMEApplication(fil.read(), _subtype = ext)
        attachedfile.add_header('content-disposition', 'attachment', filename=basename(f) )
    msg.attach(attachedfile)
    
    server.sendmail(FROM, email, msg.as_string())
    server.quit()

def apply(password, name, email, position, desc):
    addToLatex(name, desc)
    shortname = name.replace(" ", "")
    os.system("pdflatex -jobname=" + shortname + " -output-directory=tex/ tex/main.tex")

    subject = "Ansøgning til stilingen som " + position
    text = "Hej " + name + "\nJeg vedlægger min ansøgning til stillingen som " + position + ".\nMvh Morten Rasmussen"

    sendEmail(password, email, subject, text, name + ".pdf")
    cleanUpLatex(name)

def appendSent(name, email, posistion):
    f = open("sent.csv", "a")
    f.write(name + ", " + email + ", " + posistion + ", " + now.strftime("%Y-%m-%d") + "\n")
    f.close() 

def cleanUpResipients():
    f = open("recipients.csv", "w")
    f.write("Name, Email, Position, Additional Description\n")
    f.close()

def loadRecipients(password):
    with open('recipients.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                line_count += 1
            else:
                apply(password, row[0].strip(), row[1].strip(), row[2].strip(), row[3].strip())
                print("Sent to " + row[0])
                appendSent(row[0].strip(), row[1].strip(), row[2].strip())
                line_count += 1
    cleanUpResipients()
    
password = getpass.getpass("Password: ")
loadRecipients(password)