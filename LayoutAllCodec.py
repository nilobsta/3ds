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
classeur = xlrd.open_workbook(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\Codec.xls')

# Récupération du nom de toutes les feuilles sous forme de liste
sheet = classeur.sheet_names()



# Récupération de la première feuille
sheet1 = classeur.sheet_by_name(sheet[0])

#on ajoute la sheet dans le classeur de result
result1= result.add_sheet('ResultLayout')
#
for i in range(14):
     for j in range (2):
          result1.write(i,j,sheet1.cell_value(i,j))





#Connexion en SSH

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

testing = [3,4,5]




Issue = 0
Tested = 0
for i in range (14):
#for i in testing:
    Tested=Tested+1
    print(sheet1.cell_value(i,0))
    print(sheet1.cell_value(i,1))
    COMP = sheet1.cell_value(i,1) #site distant
    #Password = sheet1.cell_value(1,2) #mdp site distant

    data = [
        ('value', 'True'),
    ]
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(COMP, port=22, username="amx", password="amxcisco", allow_agent = False, look_for_keys=True) #Connexion site distant
        channel = ssh.invoke_shell()
        channel.recv(1024)
        channel.send("xCommand Presentation Start PresentationSource:4\n")
        result1.write(i,3,'Right Layout')
        channel.close()


    except: #Si la connexion SSH n'est pas faite
        result1.write(i,3,'Not Reachable')
        Issue = Issue+1

    ssh.close()






if os.path.isfile(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\LayoutResult.xls'):
    os.remove(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\LayoutResult.xls')
    result.save(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\LayoutResult.xls')
else:
    result.save(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\LayoutResult.xls')

#Création de l'objet Mail
msg = MIMEMultipart() #le mail sera stocké dans la variable sera msg

fp = open(r'C:\Users\svc_it.audiovisual\Desktop\Script\RebootMetal\LayoutResult.xls', 'rb')

file1=email.mime.base.MIMEBase('application','vnd.ms-excel')
file1.set_payload(fp.read())
fp.close()
email.encoders.encode_base64(file1)
file1.add_header('Content-Disposition','attachment;filename=CodecLayoutResult.xls')




msg['Subject'] = "Layout Codec Report"
if Issue == 0:
    body = "All CKS are rebooting (" + str(Tested) + "Tested)"
else:
    body = "Results : " + str(Issue) + " layout over " + str(Tested) + " have trouble. Please look at the file attached"


msg.attach(file1)

body=MIMEText(body)
msg.attach(body)

#Envoi mail
fromaddr = "IT.Audiovisual@noreply.ds.com" #création adresse fictive d'envoi
toaddr = "lle10@3ds.com" #destinataires
msg['From'] = "ITAudiovisual".format(fromaddr)
msg['To'] = ', '.join(toaddr)

server = smtplib.SMTP('mailhost.emea.corp.ds', 25) #serveur SMTP de destination

text = msg.as_string()
server.sendmail(fromaddr, toaddr, text) # envoi du mail

#fermeture server SMTP
server.quit()
ssh.close()



