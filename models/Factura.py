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

class Factura(object):
    def LeerCorreosFactura(self, conexion, enviar):
        print("\033[93mEMPEZAMOS EN FACTURA\033[0m")
        cursor = conexion.cursor()
        global counteraux
        def setReverseCC():
            global dicReverseCC, dicEmpreCC
            stSQL = "Select e.empresa_codigocontable, cc.cencos_digito, cc.cencos_codigo from tb_centrocosto cc INNER JOIN tb_empresa e ON cc.empresa_codigo = e.empresa_codigo where e.empresa_codigo NOT IN (11,3,6)"
            try:
                cursor.execute(stSQL)
                Msg = "Consulta de la CC " + str(stSQL) + " ejecutada con éxito"
                print(Msg)
            except psycopg2.Error as e:
                print(e.pgerror)
                return False
            ContAux = 0
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
            Subj = str(Subject.upper())  # se pasa todo a MAYUSCULAS
            global actionSubj, documentoSubj, ccDocumentoSubj, codCcDocumentoSubj, imgType, dicReverseCC, Emp
            m = re.search('(E[0-9]*)[;]', Subj)
            empresa = re.search('(MCT|FERRICAR|MARKETING)', Subj)
            print (email_message['Subject'])
            if m and empresa:
                Subj = m.group(1)
                Emp = empresa.group(1)
                prefijo = Subj[0:4]
                consec = Subj[4:]
                codCcDocumentoSubj = consec.lstrip("0")
                lsCampos = prefijo.split("E")
                ccDocumentoSubj = str(lsCampos[1])
                ccDocumentoSubj = str(empresaCodCont[Emp]) + str(lsCampos[1])
            else:
                mensaje = MIMEText("""El Codigo de factura del correo no es correcto: """ + Subj + """ \n\n"""
                                + str(email_message['Subject']) +
                                """
                                \nTenga en cuenta que debe tener uno de los siguientes formatos, cambiando el numero de remesa correspondiente:\n
                                E061 01\n
                                \nCordialmente,\n\n\nSistemas MCT. """)
                mensaje['From'] = "notificaciones@mct.com.co"
                mensaje['To'] = str(emailFrom)
                mensaje['Subject'] = "ERROR " + str(email_message['Subject'])
                enviar.enviar_correo("notificaciones@mct.com.co", str(emailFrom), mensaje.as_string())
                return False
            if ccDocumentoSubj and codCcDocumentoSubj and len(ccDocumentoSubj) == 5:
                ccDocumentoSubj = dicReverseCC[ccDocumentoSubj]
                return True
            else:
                mensaje = MIMEText("""Codigo de centro de costo inexistente:\n\n"""
                                + str(email_message['Subject']) +
                                """
                                \nCordialmente,\n\n\nSistemas MCT. """)
                mensaje['From'] = "notificaciones@mct.com.co"
                mensaje['To'] = str(emailFrom)
                mensaje['Subject'] = "ERROR " + str(email_message['Subject'])
                enviar.enviar_correo("notificaciones@mct.com.co", str(emailFrom), mensaje.as_string())
                return False

        def getDatosFactura(ccFactura, codCcFactura):
            print (ccFactura)
            print (codCcFactura)
            global detach_dir, detach_dir1, codigo_factura, cencos_factura, fecha_factura, ano, mes, dia, facturaSubj
            FacturaSQL = "Select factura_codigo, cencos_codigo, factura_fechacreacion from tb_factura where cencos_codigo = '" + str(ccFactura) + "' AND factura_numerodocumento = '" + str(codCcFactura) + "'"
            Msg = "Consulta de la Factura CC " + str(ccFactura) + " DigCCResolucion " + str(codCcFactura) + " ejecutada con éxito"
            try:
                print (FacturaSQL)
                cursor.execute(FacturaSQL)
                print (Msg)
                outfile = '.'
                ContAux = 0
                for row in cursor.fetchall():
                    print ("paso a consultar")
                    r = reg(cursor, row)
                    facturaSubj = r.factura_codigo
                    codigo_factura = r.factura_codigo
                    cencos_factura = r.cencos_codigo
                    fecha_factura = str(r.factura_fechacreacion)
                    ano = fecha_factura.split("-")[0]
                    mes = fecha_factura.split("-")[1]
                    dia = fecha_factura.split("-")[2]
                    detach_dir = '/documents/facturas/' + str(Emp) + '/' + str(ano) + '/' + str(cencos_factura) + '/' + str(mes) + '/'
                    detach_dir1 = '/usr/local/apache/htdocs/switrans/images/facturas/' + str(Emp) + '/' + str(ano) + '/' + str(
                        cencos_factura) + '/' + str(mes) + '/'
                    ContAux = ContAux + 1
                if ContAux > 0:
                    return True
                else:
                    return False
            except psycopg2.Error as e:
                print (e.pgerror)
                return False
        
        def insertDigitalizado():
            global counter
            errorFile = 0
            modulo = None
            cod_documento = None
            for part in email_message.walk():
                print ("Tipo Documento: ")
                print (part.get_content_type())
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
                print (file_type)

                filename = part.get_filename()
                if not filename:
                    filename = name_pat.findall(part.get('Content-Type'))[0][6:-1]
                if filename.endswith('.pdf'):
                    continue

                modulo = 120
                cod_documento = codigo_factura
                imgType = 122
                usuario_codigo = 1
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
                # download the attachments from email to the designated directory
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
                        print ("error")
                if not os.path.isfile(att_path):
                    # Almacenando el archivo
                    fp = open(att_path, 'wb')
                    temp = part.get_payload(decode=True)
                    fp.write(temp)
                    fp.close()
                sizeBytes = os.stat(att_path).st_size
                Msg = "Factura Almacenada en: " + str(att_path)
                print (Msg)
                Msg = "Peso en Disco Final: " + str(sizeBytes / 1024) + ' KB'
                print (Msg)
                print(detach_dir1)
                InsertFactura = (
                        "INSERT INTO tb_digitalizado "
                        "(modulo_codigo,digitalizado_documento,imadoc_codigo,digitalizado_archivo,digitalizado_ruta,digitalizado_fechacreacion,"
                        "digitalizado_fechamodificacion,usuario_codigo,digitalizado_observaciones,digitalizado_medio) "
                        "VALUES "
                        "(" + str(modulo) + ",'" + str(cod_documento) + "'," + str(imgType) + ",'" + filename + "','" + str(
                    detach_dir1) + "','" + str(datenow) + "','" + str(datenow) + "'," + str(
                    usuario_codigo) + ",'','CORREO'  ) "
                )
                try:
                    cursor.execute(InsertFactura)
                    conexion.commit()
                    Msg = "Factura cargada a base de datos tb_digitalizado\n._._._._._._._._._._._._._"
                    print (Msg)
                    # except Exception, e:
                except psycopg2.Error as e:
                    Msg = (e.pgerror)
                    print (Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(InsertFactura))
                    errorFile = 1

                UpdateFactura = (
                    "UPDATE tb_factura SET archivo_facturacionelectronica = '" + filename + "' WHERE factura_codigo = " + str(cod_documento)
                )
                try:
                    cursor.execute(UpdateFactura)
                    conexion.commit()
                    Msg = "Factura Actualizada a base de datos tb_factura, campo: archivo_facturacionelectronica \n._._._._._._._._._._._._._"
                    print (Msg)
                    # except Exception, e:
                except psycopg2.Error as e:
                    Msg = (e.pgerror)
                    print (Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(UpdateFactura))
                    errorFile = 1

                msg_factura = "Se carga la factura: " + str(cod_documento) + " a la plataforma Switrans"
                InsertFacturaEvento = (
                        "INSERT INTO tb_eventodocumento "
                        "(evedoc_documento,evedoc_relacion,evedoc_evento,modulo_codigo,evedoc_fechacreacion,evedoc_horacreacion,"
                        "usuario_codigo) "
                        "VALUES "
                        "("+ "'" + str(cod_documento) + "','" + str(cod_documento) + "','" + str(msg_factura) + "'," + str(modulo) + ",'" + str(datenow) + "','" + str(
                    hournow) + "'," + str(usuario_codigo) + ")")

                try:
                    cursor.execute(InsertFacturaEvento)
                    conexion.commit()
                    Msg = "Evento cargado a base de datos tb_eventodocumento\n._._._._._._._._._._._._._"
                    print (Msg)
                except psycopg2.Error as e:
                    Msg = e.pgerror
                    print (Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(InsertFacturaEvento))
                    errorFile = 1
                # se aumenta contador para alimentar semilla del nombre de archivo
                counter = counter + 1
                if errorFile == 1:
                    enviar.correoError(cod_documento, filename, Msg, modulo, emailFrom, email_message["From"])
                else:
                    enviar.correoConfirmacion(cod_documento, modulo, emailFrom, email_message["Subject"])

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
