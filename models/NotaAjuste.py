#! /usr/bin/env python
# -*- coding: utf-8 -*-
import imaplib, email
from email.header import decode_header
import HTMLParser
import os
import psycopg2
import psycopg2.extras
import psycopg2.extensions
from models.Reg import reg
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)
import datetime, re
from configparser import ConfigParser
import email
from email.MIMEText import MIMEText

Emp = ""
# emailFrom = "doc_soporte@mct.com.co"
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

class NotaAjuste(object):
    def LeerCorreosNotaAjuste(self, conexion, enviar):
        print("\033[93mEMPEZAMOS EN NOTA AJUSTE\033[0m")
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
            Subj = str(Subject.upper())
            global actionSubj, ccNotaSubj, codCcNotaSubj, imgType, dicReverseCC, Emp, empresaCodigo
            m = re.search('(N[0-9]*)[;]', Subj)
            empresa = re.search('(MCT|FERRICAR|MARKETING)', Subj)
            print (email_message['Subject'])
            if m and empresa:
                Subj = m.group(1)
                Emp = empresa.group(1)
                prefijo = Subj[0:4]
                consec = Subj[4:]
                codCcNotaSubj = consec.lstrip("0")
                lsCampos = prefijo.split("N")
                ccDocumento = str(lsCampos[1])
                ccNotaSubj = str(empresaCodCont[Emp]) + ccDocumento
                empresaCodigo = empresaCodGeneral[Emp]
                print (empresaCodigo)
            else:
                mensaje = MIMEText("""El Codigo de la Nota Ajuste no es correcto: """ + Subj + """ \n\n"""
                                + str(email_message['Subject']) +
                                """
                                \nTenga en cuenta que debe tener uno de los siguientes formatos, cambiando el numero de documento correspondiente:\n
                                N061 01\n
                                \nCordialmente,\n\n\nSistemas MCT. """)
                mensaje['From'] = "notificaciones@mct.com.co"
                mensaje['To'] = str(emailFrom)
                mensaje['Subject'] = "ERROR " + str(email_message['Subject'])
                enviar.enviar_correo("notificaciones@mct.com.co", str(emailFrom), mensaje.as_string())
                return False
            if ccNotaSubj and codCcNotaSubj and len(ccNotaSubj) == 5:
                ccNotaSubj = dicReverseCC[ccNotaSubj]
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
        
        def getDatosNotaAjuste(empCodigo, codCcNotaAjuste, ccNotaSubj):
            print (empCodigo)
            print (codCcNotaAjuste)
            global detach_dir, detach_dir1, notaju_codigo, cencos_docsop, fecha_docsop, ano, mes, dia
            docsopSQL = "SELECT na.notaju_codigo, na.notaju_fechacreacion, na. docsop_codigo, na.empresa_codigo, na.cencos_codigo FROM tb_notaajuste na WHERE na.empresa_codigo = " + str(empCodigo) + " AND na.notaju_cencoscodigo = "+ str(codCcNotaAjuste)
            Msg = "Consulta de la Nota Ajuste de la Empresa " + str(empCodigo) + " DigCCResolucion " + str(codCcNotaAjuste) + " ejecutada con éxito"
            try:
                print (docsopSQL)
                cursor.execute(docsopSQL)
                print (Msg)
                outfile = '.'
                ContAux = 0
                for row in cursor.fetchall():
                    print ("paso a consultar")
                    r = reg(cursor, row)
                    notaju_codigo = r.notaju_codigo
                    cencos_docsop = r.cencos_codigo
                    fecha_docsop = str(r.notaju_fechacreacion)
                    ano = fecha_docsop.split("-")[0]
                    mes = fecha_docsop.split("-")[1]
                    dia = fecha_docsop.split("-")[2]
                    print ("##########################")
                    print( notaju_codigo)
                    detach_dir = '/documents/documentosoporte/' + str(Emp) + '/' + str(ano) + '/' + str(cencos_docsop) + '/' + str(
                        mes) + '/'
                    detach_dir1 = '/usr/local/apache/htdocs/switrans/images/documentosoporte/' + str(Emp) + '/' + str(
                        ano) + '/' + str(cencos_docsop) + '/' + str(mes) + '/'
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
                print ("Nota Ajuste: ")
                print( part.get_content_type())
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
                modulo = 346
                cod_documento = notaju_codigo
                imgType = 285
                usuario_codigo = 1
                sqlDigitalizado = "SELECT digitalizado_documento, digitalizado_archivo FROM tb_digitalizado WHERE modulo_codigo = " + str(
                    modulo) + " AND  imadoc_codigo = " + str(imgType) + " AND digitalizado_documento = " + str(cod_documento)
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
                    if (existeDigitalizado):
                        msg = "El documento: " + str(digitalizado_documento) + ", ya cuenta con archivo .zip"
                        print(msg)
                        enviar.correoError(cod_documento, filename, msg, modulo, emailFrom, email_message["Subject"])
                        break
                except psycopg2.Error as e:
                    print(e.pgerror)
                    return False

                outfile = 'Min_' + str(filename)
                print("semana " + str(filename) + " counter: " + str(counter))
                # att_path = os.path.join(detach_dir, filename)
                # minfile = os.path.join(detach_dir, outfile)
                ArchivoConda = "/home/orlin/Descargas/"
                att_path = os.path.join(ArchivoConda, filename)
                minfile = os.path.join(ArchivoConda, outfile)
                digitalizado_ruta = os.path.join(detach_dir1, filename)
                print("Voy a gurdar el ftp: "+digitalizado_ruta)
                if not os.path.exists(os.path.dirname(att_path)):
                    try:
                        os.makedirs(os.path.dirname(att_path))
                    except OSError as exc:
                        print ("error al momento de crear..............")
                if not os.path.isfile(att_path):
                    fp = open(att_path, 'wb')
                    temp = part.get_payload(decode=True)
                    fp.write(temp)
                    fp.close()
                sizeBytes = os.stat(att_path).st_size
                Msg = "Nota Ajuste Almacenado en: " + str(att_path)
                print (Msg)
                Msg = "Peso en Disco Final: " + str(sizeBytes / 1024) + ' KB'
                print (Msg)
                InsertNota = (
                        "INSERT INTO tb_digitalizado "
                        "(modulo_codigo,digitalizado_documento,imadoc_codigo,digitalizado_archivo,digitalizado_ruta,digitalizado_fechacreacion,"
                        "digitalizado_fechamodificacion,usuario_codigo,digitalizado_observaciones,digitalizado_medio) "
                        "VALUES "
                        "(" + str(modulo) + ",'" + str(cod_documento) + "'," + str(imgType) + ",'" + filename + "','" + str(
                    detach_dir1) + "','" + str(datenow) + "','" + str(datenow) + "'," + str(
                    usuario_codigo) + ",'','CORREO'  ) "
                )

                try:
                    cursor.execute(InsertNota)
                    conexion.commit()
                    Msg = "Nota Ajuste cargada a base de datos tb_digitalizado\n._._._._._._._._._._._._._"
                    print(Msg)
                except psycopg2.Error as e:
                    Msg = e.pgerror
                    print (Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(InsertNota))
                    errorFile = 1

                UpdateNota = (
                            "UPDATE tb_notaajuste SET archivo_notaajuste_electronico = '" + filename + "' WHERE notaju_codigo = " + str(
                        cod_documento))
                
                try:
                    cursor.execute(UpdateNota)
                    conexion.commit()
                    Msg = "Nota Ajuste Actualizada a base de datos tb_notaajuste, campo: archivo_notaajuste_electronico \n._._._._._._._._._._._._._"
                    print(Msg)
                except psycopg2.Error as e:
                    Msg = e.pgerror
                    print (Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(UpdateNota))
                    errorFile = 1
                msg_nota = "Se carga la nota ajuste: " + str(cod_documento) + " a la plataforma Switrans"
                InsertDocsopEvento = (
                        "INSERT INTO tb_eventodocumento "
                        "(evedoc_documento,evedoc_relacion,evedoc_evento,modulo_codigo,evedoc_fechacreacion,evedoc_horacreacion,"
                        "usuario_codigo) "
                        "VALUES "
                        "(" + "'" + str(cod_documento) + "','" + str(cod_documento) + "','" + str(msg_nota) + "'," + str(
                    modulo) + ",'" + str(datenow) + "','" + str(
                    hournow) + "'," + str(usuario_codigo) + ")")

                try:
                    cursor.execute(InsertDocsopEvento)
                    conexion.commit()
                    Msg = "Evento cargado a base de datos tb_eventodocumento\n._._._._._._._._._._._._._"
                    print (Msg)
                except psycopg2.Error as e:
                    Msg = e.pgerror
                    print (Msg)
                    print("ERROR FATAL EN LA CONSULTA " + str(InsertDocsopEvento))
                    errorFile = 1
                counter = counter + 1
                
                if errorFile == 1:
                    enviar.correoError(cod_documento, filename, Msg, modulo, emailFrom, email_message["Subject"])
                else:
                    enviar.correoConfirmacion(cod_documento, modulo, emailFrom, email_message["Subject"])
        
        M = imaplib.IMAP4_SSL(config.get('conf','IMAP_HOST'), config.getint('conf','IMAP_PORT'))
        M.login(config.get('conf','CORREO_O'), config.get('conf','PASS'))
        M.select(config.get('conf','NA'))
        # M.login(config.get('conf','IMAP_CORREO_DA'), config.get('conf','IMAP_PASSWORD_DA'))
        # M.select(config.get('conf','IMAP_NA'))
        name_pat = re.compile('name=".*"')
        Msg = "\033[92mListo para traer mensajes sin leer de la cuenta doc_soporte@mct.com.co NOTA AJUSTE\033[0m"
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
                # eMail = re.findall(r'\W*(fe)\W*@[\w\.-]+', email_message['From'])
                # if (len(eMail) <= 0):
                #     Msg = "El Correo es enviado desde un contacto NO valido debe ser fe@mct.com.co"
                #     print (Msg)
                #     break
                Subj = decode_header(email_message['Subject'])[0][0]
                if emailSubjectCheck(Subj) is False:
                    Msg = "El Asunto no es correcto o esta mal formado (no tiene el numero de documento soporte), se envía correo y se detiene el proceso"
                    print (Msg)
                    break
                else:
                    print ("Cencos: " + str(ccNotaSubj))
                    print ("cencoscodigo: " + str(codCcNotaSubj))
                    print ("empresacodigo: " + str(empresaCodigo))
                    if getDatosNotaAjuste(empresaCodigo, codCcNotaSubj, ccNotaSubj) is False:
                        Msg = "No se lograron obtener datos del Documento Soporte"
                        print (Msg)
                        break
                    insertDigitalizado()
            counteraux = counteraux + 1
        if counter == 0:
            Msg = "No existen mensajes sin leer en la cuenta"
            print (Msg)
        M.close()
        M.logout

        Msg = "\033[93m*********************************************** Proceso Finalizado NOTA AJUSTE ***********************************************\033[0m\n\n"
        print (Msg)