#! /usr/bin/env python
# -*- coding: utf-8 -*-
from configparser import ConfigParser
import imaplib, email, base64
from email.header import decode_header
import smtplib
from conexiones.EnviarCorreo import EnviarCorreo
import HTMLParser
import os, sys
import psycopg2
import psycopg2.extras
import psycopg2.extensions

from conexiones.ConexionDB import Conexion
from models.Reg import reg
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)
import time, math, array
from glob import glob
import datetime, re
from PIL import Image

from smtplib import SMTP  # this invokes the secure SMTP protocol (port 465, uses SSL)
import email
from email.MIMEText import MIMEText

# Fecha y hora actual
dateNow = datetime.datetime.now()
hourNow = str(dateNow.hour) + ":" + str(dateNow.minute) + ":" + str(dateNow.second)

Emp = ""
# emailFrom = "facturacion@mct.com.co"
emailFrom = "orlin.carrillo@mct.com.co"
counter = 0
counterAux = 1
imgType = 0
usuarioCodigo = 0
detach_dir = '.'
detach_dir1 = '.'
dicReverseCC = {}
filename = None
outfile = None
ano = 0
mes = 0
dia = 0
Msg = None
codigoNotaDebito = None
cencosNotaDebito = None
fechaNotaDebito = None
empresaCodCont = {
  "MCT": "01",
  "FERRICAR": "13",
  "MARKETING": "02"
}
empresaCodGeneral = {
  "01": "1",
  "13": "12",
  "02": "2"
}

parser = HTMLParser.HTMLParser()
datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

config = ConfigParser()
config.read('config/config.ini')

class NotaDebito(object):
    def LeerCorreosNotaDebito(self, conexion, enviar):
        print("\033[93mEMPEZAMOS EN NOTA DEBITO\033[0m")
        cursor = conexion.cursor() 
        global counteraux, counterAux
        def setReverseCC():
            global dicReverseCC, dicEmpreCC
            stSQL = "SELECT e.empresa_codigocontable, cc.cencos_digito, cc.cencos_codigo FROM tb_centrocosto cc INNER JOIN tb_empresa e ON cc.empresa_codigo = e.empresa_codigo where e.empresa_codigo NOT IN (11,3,6)"
            try:
                cursor.execute(stSQL)
                Msg = "Consulta de la CC " + str(stSQL) + " ejecutada con exito"
                print(Msg)
            except psycopg2.Error as e:
                print(e.pgerror)
                return False
            
            for row in cursor.fetchall():
                r = reg(cursor, row)
                stDigCC = str(r.cencos_digito).strip()
                if len(stDigCC) < 3:
                    for x in range(len(stDigCC) + 1, 4):
                        stDigCC = "0" + str(stDigCC)
                stPrefCCReverse = str(r.empresa_codigocontable) + str(stDigCC)
                dicReverseCC[stPrefCCReverse] = r.cencos_codigo
        setReverseCC()
        
        def emailSubjectCheck(Subject):
            global imgType, dicReverseCC, Emp, consecutivo, codigoEmpresa
            Subj = str(Subject.upper())  # se pasa todo a MAYUSCULAS
            m = re.search('(ND[0-9]*)[;]', Subj)
            empresa = re.search('(MCT|FERRICAR|MARKETING)', Subj)
            print ("Este es el subject del emailSubjectCheck: " + email_message['Subject'])
            if m and empresa:
                Subj = m.group(1)
                print ("Group: " + Subj)
                Emp = empresa.group(1)
                print ("Emp: " + Emp)
                consecutivo = Subj.split("ND")[1]
                codigoEmpresa = empresaCodCont[Emp]
                codigoEmpresa = empresaCodGeneral[codigoEmpresa]
                print ("Consecutivo: " + consecutivo)
                print ("codigoEmpresa: " + codigoEmpresa)
            else:
                mensaje = MIMEText("""El Codigo de la nota debito no es correcto: """ + Subj + """ \n\n"""
                                + str(email_message['Subject']) +
                                """
                                \nTenga en cuenta que debe tener uno de los siguientes formatos, cambiando el numero de remesa correspondiente:\n
                                E061 01\n
                                \nCordialmente,\n\n\nSistemas MCT. """)
                mensaje['From'] = "notificaciones@mct.com.co"
                mensaje['To'] = str(emailFrom)
                mensaje['Subject'] = "ERROR " + str(email_message['Subject'])
                enviar.enviar_correo("notificaciones@mct.com.co",
                            str(emailFrom),
                            mensaje.as_string())
                return False
            return True
        
        def getDatosNotaDebito(empCodigo, consecutivoCC):
            global detach_dir, detach_dir1, codigoNotaDebito, cencosNotaDebito, fechaNotaDebito, ano, mes, dia
            NotaDebitoSQL = "SELECT notadebito_codigo, nd.cencos_codigo, nd.notadebito_fechacreacion FROM tb_notadebito nd LEFT JOIN tb_centrocosto cc ON nd.cencos_codigo = cc.cencos_codigo WHERE cc.empresa_codigo = " + str(empCodigo) + " AND nd.notadebito_cencoscodigo = " + str(consecutivoCC)
            Msg = "Consulta de la nota debito para la empresa " + str(empCodigo) + " ccNotaDebito " + str(consecutivoCC) + " ejecutada con exito"
            try:
                print (NotaDebitoSQL)
                cursor.execute(NotaDebitoSQL)
                print (Msg)
                outfile = '.'
                ContAux = 0
                for row in cursor.fetchall():
                    print("paso a consultar")
                    r = reg(cursor, row)
                    codigoNotaDebito = r.notadebito_codigo
                    cencosNotaDebito = r.cencos_codigo
                    fechaNotaDebito = str(r.notadebito_fechacreacion)
                    ano = fechaNotaDebito.split("-")[0]
                    mes = fechaNotaDebito.split("-")[1]
                    dia = fechaNotaDebito.split("-")[2]
                    detach_dir = '/documents/facturas/' + str(Emp) + '/' + str(ano) + '/' + str(cencosNotaDebito) + '/' + str(mes) + '/'
                    detach_dir1 = '/usr/local/apache/htdocs/switrans/images/facturas/' + str(Emp) + '/' + str(ano) + '/' + str(
                        cencosNotaDebito) + '/' + str(mes) + '/'
                    ContAux = ContAux + 1        
                if ContAux > 0:
                    return True
                else:
                    return False
            except psycopg2.Error as e:
                print(e.pgerror)
                return False

        def insertDigitalizado():
            global counter
            errorFile = 0
            modulo = None
            cod_documento = None
            for part in email_message.walk():
                print(part.get_content_type())
                if part.is_multipart():
                    continue
                if part.get_content_type() == 'text/plain':
                    body = "\n" + part.get_payload() + "\n"
                if part.get_content_maintype() == 'text':
                    continue
                if part.get_content_type() == 'application/xml':
                    continue
                if part.get_content_type() == 'application/pdf':
                    continue
                if part.get_content_type() == 'multipart/mixed':
                    continue
                
                file_type = part.get_content_type().split('/')[1]
                print(file_type)

                filename = part.get_filename()
                if not filename:
                    filename = name_pat.findall(part.get('Content-Type'))[0][6:-1]
                if filename.endswith('.pdf'):
                    continue
                modulo = 256
                cod_documento = codigoNotaDebito
                imgType = 140
                usuarioCodigo = 1
                outfile = 'Min_' + str(filename)
                # Validacion de existencia de archivo .zip
                sqlDigitalizado = "SELECT digitalizado_documento, digitalizado_archivo FROM tb_digitalizado WHERE modulo_codigo = "+str(modulo)+" AND  imadoc_codigo = "+str(imgType)+" AND digitalizado_documento = "+str(cod_documento)
                digitalizado_documento = ""
                digitalizado_archivo = ""
                existeDigitalizado = False
                try:
                    print(sqlDigitalizado)
                    cursor.execute(sqlDigitalizado)
                    for row in cursor.fetchall():
                        r = reg(cursor, row)
                        digitalizado_documento = r.digitalizado_documento
                        digitalizado_archivo = r.digitalizado_archivo
                    existeDigitalizado = (".zip" in digitalizado_archivo)
                    if(existeDigitalizado):
                        msg = "El documento: " +str(digitalizado_documento)+ ", ya cuenta con archivo .zip"
                        print(msg)
                        enviar.correoError(cod_documento, filename, msg, modulo, emailFrom, email_message["Subject"])
                        break
                except psycopg2.Error as e:
                    print (e.pgerror)
                    return False
                
                outfile = 'Min_' + str(filename)
                print("semana " + str(filename) + " counter: " + str(counter))
                # att_path = os.path.join(detach_dir, filename)
                # minfile = os.path.join(detach_dir, outfile)
                ArchivoConda = "/home/orlin/Descargas/"
                att_path = os.path.join(ArchivoConda, filename)
                minfile = os.path.join(ArchivoConda, outfile)
                digitalizado_ruta = os.path.join(detach_dir1, filename)
                if not os.path.exists(os.path.dirname(att_path)):
                    try:
                        os.makedirs(os.path.dirname(att_path))
                    except OSError as exc:
                        print("error")
                if not os.path.isfile(att_path):
                    # Almacenando el archivo
                    fp = open(att_path, 'wb')
                    temp = part.get_payload(decode=True)
                    fp.write(temp)
                    fp.close()
                sizeBytes = os.stat(att_path).st_size
                Msg = "Nota Debito Almacenada en: " + str(att_path)
                print(Msg)
                Msg = "Peso en Disco Final: " + str(sizeBytes / 1024) + ' KB'
                print(Msg)
                print(codigoNotaDebito)
                InsertFactura = (
                        "INSERT INTO tb_digitalizado "
                        "(modulo_codigo,digitalizado_documento,imadoc_codigo,digitalizado_archivo,digitalizado_ruta,digitalizado_fechacreacion,"
                        "digitalizado_fechamodificacion,usuario_codigo,digitalizado_observaciones,digitalizado_medio) "
                        "VALUES "
                        "(" + str(modulo) + ",'" + str(codigoNotaDebito) + "'," + str(imgType) + ",'" + filename + "','" + str(
                    detach_dir1) + "','" + str(dateNow) + "','" + str(dateNow) + "'," + str(
                    usuarioCodigo) + ",'','CORREO'  ) "
                )
                try:
                    cursor.execute(InsertFactura)
                    conexion.commit()
                    Msg = "Nota debito cargada a base de datos tb_digitalizado\n._._._._._._._._._._._._._"
                    print(Msg)
                    # except Exception, e:
                except psycopg2.Error as e:
                    Msg = e.pgerror
                    print(Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(InsertFactura))
                    errorFile = 1

                UpdateFactura = (
                    "UPDATE tb_notacontabilidad SET archivo_facturacionelectronica = '" + filename + "' WHERE notcon_codigo = " + str(cod_documento)
                )
                try:
                    cursor.execute(UpdateFactura)
                    conexion.commit()
                    Msg = "Nota debito Actualizada a base de datos tb_notacontabilidad, campo: archivo_facturacionelectronica \n._._._._._._._._._._._._._"
                    print(Msg)
                    # except Exception, e:
                except psycopg2.Error as e:
                    Msg = e.pgerror
                    print(Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(UpdateFactura))
                    errorFile = 1

                msg_notaDebito = "Se carga la nota Debito : " + str(cod_documento) + " a la plataforma Switrans"
                InsertNotaEvento = (
                        "INSERT INTO tb_eventodocumento "
                        "(evedoc_documento,evedoc_relacion,evedoc_evento,modulo_codigo,evedoc_fechacreacion,evedoc_horacreacion,"
                        "usuario_codigo) "
                        "VALUES "
                        "("+ "'" + str(cod_documento) + "','" + str(cod_documento) + "','" + str(msg_notaDebito) + "'," + str(modulo) + ",'" + str(dateNow) + "','" + str(
                    hourNow) + "'," + str(usuarioCodigo) + ")")

                try:
                    cursor.execute(InsertNotaEvento)
                    conexion.commit()
                    Msg = "Evento cargado a base de datos tb_eventodocumento\n._._._._._._._._._._._._._"
                    print(Msg)
                except psycopg2.Error as e:
                    Msg = e.pgerror
                    print(Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(InsertNotaEvento))
                    errorFile = 1
                # se aumenta contador para alimentar semilla del nombre de archivo
                counter = counter + 1

                if errorFile == 1:
                    enviar.correoError(cod_documento, filename, Msg, modulo, emailFrom, email_message["Subject"])
                else:
                    enviar.correoConfirmacion(cod_documento, modulo, emailFrom, email_message["Subject"]) 
        
        M = imaplib.IMAP4_SSL(config.get('conf','IMAP_HOST'), config.getint('conf','IMAP_PORT'))
        M.login(config.get('conf','CORREO_O'), config.get('conf','PASS'))
        M.select(config.get('conf','IMAP_ND'))
        # M.login(config.get('conf','IMAP_CORREO_FCD'), config.get('conf','IMAP_PASSWORD_FCD'))
        # M.select(config.get('conf','IMAP_ND')
        name_pat = re.compile('name=".*"')
        Msg = "\033[92mListo para traer mensajes sin leer de la cuenta carga_fe@mct.com.co NOTA DEBITO\033[0m"
        print(Msg)
        typ, data = M.search(None, 'UNSEEN')
        for num in data[0].split():
            typ, data = M.fetch(num, '(RFC822)')
            if typ == 'OK':
                raw_email = data[0][1]
                raw_email_string = raw_email.decode('utf-8')
                email_message = email.message_from_string(raw_email_string)
                Msg = '\033[94mmensaje ' + str(counterAux) + ' obtenido con exito: ' + email_message['From'] + ' con subject: ' + \
                    email_message['Subject'] + '\033[0m'
                print(Msg)
                eMail = re.findall(r'\W*(fe)\W*@[\w\.-]+', email_message['From'])
                if(len(eMail) <= 0):
                    Msg = "El Correo es enviado desde un contacto NO valido debe ser fe@mct.com.co"
                    print (Msg)
                    break
                Subj = decode_header(email_message['Subject'])[0][0]
                print("Subj: " + Subj)
                if emailSubjectCheck(Subj) is False:
                    Msg = "El Asunto no es correcto o esta mal formado (no tiene el numero de factura), se envia correo y se detiene el proceso"
                    print(Msg)
                    break
                else:
                    if getDatosNotaDebito(codigoEmpresa, consecutivo) is False:
                        Msg = "No se lograron obtener datos de la Factura"
                        print(Msg)
                        break
                    insertDigitalizado()
            counterAux = counterAux + 1

        if counter == 0 :
            Msg = "No existen mensajes sin leer en la cuenta"
            print(Msg)

        M.close()
        M.logout()

        Msg = "\033[93m*********************************************** Proceso Finalizado NOTA DEBITO ***********************************************\033[0m\n\n"
        print(Msg)
        