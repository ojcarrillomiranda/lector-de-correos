#! /usr/bin/env python
# -*- coding: utf-8 -*-
import email
import imaplib
import os
import datetime
import re
import psycopg2
from email.header import decode_header
from configparser import ConfigParser
from email.mime.text import MIMEText

from models.reg import reg

empresa = ""

config = ConfigParser()
config.read('config/config.ini')
host = config.get('leer_correo','IMAP_HOST')
clave =  config.getint('leer_correo','IMAP_PORT')
email_notificaciones = config.get('enviar_correo','SMTP_CORREO')

counter = 0
counteraux = 1
imagdoc_codigo = 0
usuario_codigo = 0
detach_dir = '.'
detach_dir1 = '.'
dicc_reverse_cencos = {}
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
empresa_codigo_cont = {
  "MCT": "01",
  "FERRICAR": "13",
  "MARKETING": "02"
}
empresa_codigo_general = {
  "MCT": "1",
  "FERRICAR": "12",
  "MARKETING": "2"
}

datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

class Funcion:

    def __init__(self, cursor, conexion, enviar):
        self.conexion = conexion
        self.cursor = cursor
        self.enviar = enviar

    def get_datos_centro_costo(self):
        global dicc_reverse_cencos, dic_empresa_cencos
        stSQL = "Select e.empresa_codigocontable, cc.cencos_digito, cc.cencos_codigo from tb_centrocosto cc INNER JOIN tb_empresa e ON cc.empresa_codigo = e.empresa_codigo where e.empresa_codigo NOT IN (11,3,6)"
        try:
            self.cursor.execute(stSQL)
            Msg = "Consulta de la CC set_reverse_cc " + str(stSQL) + " ejecutada con éxito"
            print(Msg+"\n")
        except psycopg2.Error as e:
            print(e.pgerror)
            return False
        ContAux = 0
        for row in self.cursor.fetchall():
            r = reg(self.cursor, row)
            stDigCC = str(r.cencos_digito).strip()
            if len(stDigCC) < 3:
                for x in range(len(stDigCC) + 1, 4):
                    stDigCC = "0" + str(stDigCC)
            stPrefCCReverse = str(r.empresa_codigocontable) + str(stDigCC)
            dicc_reverse_cencos[stPrefCCReverse] = r.cencos_codigo

    def validar_asunto_correo(self, letra, asunto, email_from):
        global doc_asunto, cencos_doc_asunto, cod_doc_asunto, imagdoc_codigo, empresa,empresa_codigo,codigo_empresa,consecutivo
        Subj = str(asunto.upper())  # se pasa todo a MAYUSCULAS
        m = re.search('('+letra+'[0-9]*)[;]', Subj)
        empresa = re.search('(MCT|FERRICAR|MARKETING)', Subj)
        if m and empresa:
            Subj = m.group(1)
            empresa = empresa.group(1)
            prefijo = Subj[0:4]
            consec = Subj[4:]
            cod_doc_asunto = consec.lstrip("0")
            ls_campos = prefijo.split(letra)
            consecutivo = Subj.split(letra)[1]
            cencos_doc = str(ls_campos[1])
            empresa_codigo = empresa_codigo_general[empresa]
            if (cencos_doc == "130"):  # Cambio centro de costo para el caso de Ferricar
                cencos_doc = "030"
            empresa_codigo_contable = empresa_codigo_cont[empresa]
            cencos_doc_asunto = str(empresa_codigo_cont[empresa]) + cencos_doc
        else:
            mensaje = MIMEText("""El Codigo de factura del correo no es correcto: """ + asunto + """ \n\n"""
                               + str(email_message['Subject']) +
                               """
                               \nTenga en cuenta que debe tener uno de los siguientes formatos, cambiando el numero de remesa correspondiente:\n"""
                               +letra+"""061 01\n
                            \nCordialmente,\n\n\nSistemas MCT. """)
            mensaje['From'] = email_notificaciones
            mensaje['To'] = str(email_from)
            mensaje['Subject'] = "ERROR " + str(email_message['Subject'])
            self.enviar.enviar_correo(email_notificaciones, str(email_from), mensaje.as_string())
            return False
        #si el centro de costo es correcto o viene una nota debito
        if (cencos_doc_asunto and cod_doc_asunto and len(cencos_doc_asunto) == 5) or letra == "ND":
            if  letra != "ND":
                cencos_doc_asunto = dicc_reverse_cencos[cencos_doc_asunto]
            datos_asunto = {
                "empresa":empresa,
                "cencos_doc_asunto": cencos_doc_asunto,
                "empresa_codigo": empresa_codigo,
                "cencos_doc": cencos_doc,
                "cod_doc_asunto": cod_doc_asunto,
                "consecutivo":consecutivo,
                "empresa_codigo_contable":empresa_codigo_contable
            }
            return datos_asunto
        else:
            mensaje = MIMEText("""Codigo de centro de costo inexistente:\n\n"""
                               + str(email_message['Subject']) +
                               """
                               \nCordialmente,\n\n\nSistemas MCT. """)
            mensaje['From'] = email_notificaciones
            mensaje['To'] = str(email_from)
            mensaje['Subject'] = "ERROR " + str(email_message['Subject'])
            self.enviar.enviar_correo(email_notificaciones, str(email_from), mensaje.as_string())
            return False


    def insertar_digitalizado(self, modulo, imagdoc_codigo, documento, ruta, filename, cargue_string, error_string,
                              email_from):
        global counter,errorFile
        usuario_codigo = 1
        errorFile = None
        # Validacion de existencia de archivo .zip
        sql_digitalizado = "SELECT digitalizado_documento, digitalizado_archivo FROM tb_digitalizado WHERE modulo_codigo = "+str(modulo)+" AND  imadoc_codigo = "+str(imagdoc_codigo)+" AND digitalizado_documento = "+str(documento)
        digitalizado_documento = ""
        digitalizado_archivo = ""
        existeDigitalizado = False
        try:
            print("\033[93m######################### SQL VALIDACION SI EXISTE EN TB_DIGITALIZADO #################################\033[0m")
            print("\033[93m" + sql_digitalizado + "\033[0m")
            self.cursor.execute(sql_digitalizado)
            for row in  self.cursor.fetchall():
                r = reg(self.cursor, row)
                digitalizado_documento = r.digitalizado_documento
                digitalizado_archivo = r.digitalizado_archivo
            existeDigitalizado = (".zip" in digitalizado_archivo)
            if(existeDigitalizado):
                msg = "El documento: " +str(digitalizado_documento)+ ", ya cuenta con archivo .zip"
                print("\033[94m######################### RESULTADO SQL VALIDACION SI EXISTE EN TB_DIGITALIZADO  #################################\033[0m\n" + "\033[94mdigitalizado_documento = " + str(
                    digitalizado_documento) + "\n" + "digitalizado_archivo = " + str(
                    digitalizado_archivo) + "\033[0m")
                print("\033[91m" + msg + "\033[0m")
                self.enviar.correo_error(documento, filename, msg, cargue_string, error_string, email_from, email_message["Subject"])
                return
        except psycopg2.Error as e:
            print (e.pgerror)
            return False
        outfile = 'Min_' + str(filename)
        print("\033[94m######################### RESULTADO SQL VALIDACION SI EXISTE EN TB_DIGITALIZADO  #################################\033[0m")
        print("\033[94mdigitalizado_documento: "+str(documento)+"\033[0m")
        print("\033[94mdigitalizado_archivo: " + str(filename) + "\033[0m")
        print("\033[91mEl digitalizado documento y el Archivo zip no existen en DB, se procede a guardarlos en DB\033[0m\n")
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
        mensaje = "Documeto Almacenado en: " + str(att_path) + "\n"
        print (mensaje)
        insertar_documento = (
                "INSERT INTO tb_digitalizado "
                "(modulo_codigo,digitalizado_documento,imadoc_codigo,digitalizado_archivo,digitalizado_ruta,digitalizado_fechacreacion,"
                "digitalizado_fechamodificacion,usuario_codigo,digitalizado_observaciones,digitalizado_medio) "
                "VALUES "
                "(" + str(modulo) + ",'" + str(documento) + "'," + str(imagdoc_codigo) + ",'" + filename + "','" + str(ruta) + "','" + str(datenow) + "','" + str(datenow) + "'," + str(
            usuario_codigo) + ",'','CORREO'  ) ")
        print("\033[94m######################### SQL INSERT DIGITALIZADO #################################\033[0m")
        print("\033[94m" + insertar_documento + "\033[0m")
        try:
            self.cursor.execute(insertar_documento)
            self.conexion.commit()
            mensaje = "\033[91mDocumento cargado a base de datos tb_digitalizado\n._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._\033[0m"
            print (mensaje)
        except psycopg2.Error as e:
            mensaje = (e.pgerror)
            print (mensaje)
            print("ERROR FATAL EN LA CONSULTA " + str(insertar_documento))
            errorFile = 1
        # se aumenta contador para alimentar semilla del nombre de archivo
        counter = counter + 1
        if errorFile == 1:
            self.enviar.correo_error(documento, filename, mensaje, cargue_string, error_string, email_from, email_message["Subject"])
        else:
            self.enviar.correo_confirmacion(documento, cargue_string, email_from, email_message["Subject"])
            return filename

    def eventos(self, cod_documento, modulo):
        usuario_codigo = 1
        mensaje = "Se carga la factura: " + str(cod_documento) + " a la plataforma Switrans"
        insertar_documento_evento = (
                "INSERT INTO tb_eventodocumento "
                "(evedoc_documento,evedoc_relacion,evedoc_evento,modulo_codigo,evedoc_fechacreacion,evedoc_horacreacion,"
                "usuario_codigo) "
                "VALUES "
                "("+ "'" + str(cod_documento) + "','" + str(cod_documento) + "','" + str(mensaje) + "'," + str(modulo) + ",'" + str(datenow) + "','" + str(
            hournow) + "'," + str(usuario_codigo) + ")")
        try:
            self.cursor.execute(insertar_documento_evento)
            self.conexion.commit()
            mensaje = "\033[93mEvento cargado a base de datos tb_eventodocumento\n._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._\033[0m\n"
            print (mensaje)
        except psycopg2.Error as e:
            mensaje = e.pgerror
            print (mensaje)
            print("ERROR FATAL EN LA CONSULTA " + str(insertar_documento_evento))
            errorFile = 1

    def obtener_filename(self, email_message):
        global part
        for part in email_message.walk():
            if part.is_multipart():
                continue
            patron1 = re.search("text/plain|application/xml|application/pdf|multipart/mixed", part.get_content_type())
            patron2 = re.search("text", part.get_content_maintype())
            if patron1 > 0 or patron2 > 0:
                continue
            file_type = part.get_content_type().split('/')[1]
            filename = part.get_filename()
            if not filename:
                filename = name_pat.findall(part.get('Content-Type'))[0][6:-1]
            if filename.endswith('.pdf'):
                continue
            return filename

    def analizar_correo(self, correo, password, carpeta):
        global counteraux,email_message, name_pat, cambio
        data_asuntos = []
        cambio = re.sub("_", " ", carpeta)
        M = imaplib.IMAP4_SSL(host, clave)
        M.login(correo, password)
        M.select(carpeta)
        name_pat = re.compile('name=".*"')
        mensaje = "\033[92mListo para traer mensajes sin leer de la cuenta "+correo+" de la carpeta " +cambio.upper()+"\033[0m\n"
        print (mensaje)
        typ, data = M.search(None, 'UNSEEN')
        for num in data[0].split():
            typ, data = M.fetch(num, '(RFC822)')
            if typ == 'OK':
                raw_email = data[0][1]
                raw_email_string = raw_email.decode('utf-8')
                email_message = email.message_from_string(raw_email_string)
                # eMail = re.findall(r'\W*(fe)\W*@[\w\.-]+', email_message["From"])
                # if(len(eMail) <= 0):
                #     mensaje = "\033[94mEl Correo es enviado desde un contacto NO valido debe ser fe@mct.com.co\033[0m\n"
                #     print (mensaje)
                #     break
                filename = self.obtener_filename(email_message)
            asunto = decode_header(email_message['Subject'])[0][0]
            counteraux = counteraux + 1
            datos = {
                "asunto":asunto,
                "filename":filename
            }
            data_asuntos.append(datos)

        if counteraux > 1:
            return data_asuntos
        if counter == 0 :
            mensaje = "\033[91mNo existen mensajes sin leer en la cuenta "+ correo + " en la carpeta "+ cambio.upper()+"\033[0m"
            print (mensaje)
        M.close()
        M.logout()
        mensaje = "\033[93m*********************************************** Proceso Finalizado " +cambio.upper()+ " ***********************************************\033[0m\n\n"
        print (mensaje)
