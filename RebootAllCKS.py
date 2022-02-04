#!/usr/bin/env python
#-*- coding: utf-8 -*-


import os
import subprocess
import sys
import smtplib
from email.message import EmailMessage
import paramiko
import socket
import time
import errno
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email
import xlrd
from xlwt import Workbook

import urllib3
from urllib3 import exceptions
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)



#création classeur
result=Workbook()

# ouverture du classeur
classeur = xlrd.open_workbook(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\CKSMetal.xls')

# Récupération du nom de toutes les feuilles sous forme de liste
sheet = classeur.sheet_names()



# Récupération de la première feuille
sheet1 = classeur.sheet_by_name(sheet[0])

#on ajoute la sheet dans le classeur de result
result1= result.add_sheet('ResultReboot')
#
for i in range(21):
     for j in range (2):
          result1.write(i,j,sheet1.cell_value(i,j))





#Connexion en SSH

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

testing = [21]




Issue = 0
Tested = 0
for i in range (21):
#for i in testing:
    Tested=Tested+1
    print(sheet1.cell_value(i,1))
    print(sheet1.cell_value(i,2))
    COMP = sheet1.cell_value(i,2) #site distant
    #Password = sheet1.cell_value(1,2) #mdp site distant

    data = [
        ('value', 'True'),
    ]
    try:
        requests.put(COMP, data=data, verify=False, auth=('integrator', 'integrator'))
        result1.write(i,3,'CKS Rebooted')



    except: #Si la connexion SSH n'est pas faite
        result1.write(i,3,'Not Reachable')
        Issue = Issue+1

    ssh.close()






if os.path.isfile(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\CKSResult.xls'):
    os.remove(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\CKSResult.xls')
    result.save(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\CKSResult.xls')
else:
    result.save(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\CKSResult.xls')

#Création de l'objet Mail
msg = MIMEMultipart() #le mail sera stocké dans la variable sera msg

fp = open(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\CKSResult.xls', 'rb')

file1=email.mime.base.MIMEBase('application','vnd.ms-excel')
file1.set_payload(fp.read())
fp.close()
email.encoders.encode_base64(file1)
file1.add_header('Content-Disposition','attachment;filename=CKSRebootResult.xls')




msg['Subject'] = "Reboot CKS Report"
if Issue == 0:
    body = "All CKS are rebooting"
else:
    body = "Results : " + str(Issue) + " CKS over " + str(Tested) + " have trouble. Please look at the file attached"


msg.attach(file1)

body=MIMEText(body)
msg.attach(body)

#Envoi mail
fromaddr = "IT.Audiovisual@noreply.ds.com" #création adresse fictive d'envoi
toaddr = "lle10@3ds.com", "lbi1@3ds.com" #destinataires
msg['From'] = "ITAudiovisual".format(fromaddr)
msg['To'] = ', '.join(toaddr)

server = smtplib.SMTP('mailhost.emea.corp.ds', 25) #serveur SMTP de destination

text = msg.as_string()
server.sendmail(fromaddr, toaddr, text) # envoi du mail

#fermeture server SMTP
server.quit()
ssh.close()



