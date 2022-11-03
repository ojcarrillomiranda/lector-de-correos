#! /usr/bin/env python
# -*- coding: utf-8 -*-
import imaplib, email
from email.header import decode_header
import HTMLParser
import os
import psycopg2
import psycopg2.extras
import psycopg2.extensions
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)
import datetime, re
from configparser import ConfigParser
import email
from email.MIMEText import MIMEText
from models.Reg import reg

Emp = ""
# emailFrom = "facturacion@mct.com.co"
emailFrom = "orlin.carrillo@mct.com.co"
counter = 0
counteraux = 1
imgType = 0
usuario_codigo = 0
detach_dir = '.'
detach_dir1 = '.'
dicReverseCC = {}
filename = None
outfile = None
ano = 0
mes = 0
dia = 0
Msg = None
codigo_nota = None
cencos_nota = None
fecha_nota = None
notaSubj = None
empresaCodCont = {
  "MCT": "01",
  "FERRICAR": "13",
  "MARKETING": "02"
}
empresaCodGeneral = {
  "MCT": "1",
  "FERRICAR": "12",
  "MARKETING": "2"
}

parser = HTMLParser.HTMLParser()
datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)


config = ConfigParser()
config.read('config/config.ini')

class Funciones():
    
    def analizar_correo():
        M = imaplib.IMAP4_SSL(config.get('conf','IMAP_HOST'), config.getint('conf','IMAP_PORT'))
        M.login(config.get('conf','CORREO_O'), config.get('conf','PASS'))
        M.select(config.get('conf','IMAP_FAC'))
        # M.login(config.get('conf','IMAP_CORREO_FCD'), config.get('conf','IMAP_PASSWORD_FCD'))
        # M.select(config.get('conf','IMAP_FAC')
        name_pat = re.compile('name=".*"')
        Msg = "\033[92mListo para traer mensajes sin leer de la cuenta carga_fe@mct.com.co FACTURA\033[0m"
        print (Msg)
        typ, data = M.search(None, 'UNSEEN')
        for num in data[0].split():
            typ, data = M.fetch(num, '(RFC822)')
            if typ == 'OK':
                raw_email = data[0][1]
                raw_email_string = raw_email.decode('utf-8')
                email_message = email.message_from_string(raw_email_string)
                Msg = '\033[94mmensaje ' + str(counteraux) + ' obtenido con éxito: ' + email_message[
                    'From'] + ' con subject: ' + \
                    email_message['Subject'] + '\033[0m'
                print(Msg)
                eMail = re.findall(r'\W*(fe)\W*@[\w\.-]+', email_message["From"])
                print("este es email de factura:" + email_message['From'])
                if(len(eMail) <= 0):
                    Msg = "El Correo es enviado desde un contacto NO valido debe ser fe@mct.com.co"
                    print (Msg)
                    break
                Subj = decode_header(email_message['Subject'])[0][0]
                if emailSubjectCheck(Subj) is False:
                    Msg = "El Asunto no es correcto o esta mal formado (no tiene el numero de documento soporte), se envía correo y se detiene el proceso"
                    print(Msg)
                    break
                else:
                    print ("Cencos: " + str(ccDocumentoSubj))
                    print ("reso: " + str(codCcDocumentoSubj))
                    if getDatosFactura(ccDocumentoSubj, codCcDocumentoSubj) is False:
                        Msg = "No se lograron obtener datos de la Factura"
                        print (Msg)
                        break
                    insertDigitalizado()
            counteraux = counteraux + 1
        if counter == 0 :
            Msg = "No existen mensajes sin leer en la cuenta"
            print (Msg)
        M.close()
        M.logout()
        Msg = "\033[93m*********************************************** Proceso Finalizado FACTURA ***********************************************\033[0m\n\n"
        print (Msg)