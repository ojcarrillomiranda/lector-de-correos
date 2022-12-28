#! /usr/bin/env python
# -*- coding: utf-8 -*-
import email
import imaplib
import os
from email.header import decode_header
import HTMLParser
import psycopg2
import psycopg2.extensions
import psycopg2.extras
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)
import datetime
import email
import re
from configparser import ConfigParser
from email.MIMEText import MIMEText

# from controllers.Controlador import Controlador
# from models import nota_credito
from models.reg import reg

empresa = ""
# emailFrom = "facturacion@mct.com.co"
emailFrom = "orlin.carrillo@mct.com.co"
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

parser = HTMLParser.HTMLParser()
datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

config = ConfigParser()
config.read('config/config.ini')
host = config.get('conf','IMAP_HOST')
clave =  config.getint('conf','IMAP_PORT')


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
            return dicc_reverse_cencos
            

    def validar_asunto_correo(self, asunto, letra):
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
            if(cencos_doc == "130"): # Cambio centro de costo para el caso de Ferricar
                cencos_doc = "030"
            empresa_codigo = empresa_codigo_general[empresa]
            #nota debito
            # codigo_empresa = empresa_codigo_cont[empresa]
            # codigo_empresa = empresa_codigo_general[codigo_empresa]
            #fin nota debito
            cencos_doc_asunto = str(empresa_codigo_cont[empresa]) +  cencos_doc
            datos_asunto = {
                "cencos_doc_asunto":cencos_doc_asunto,
                "empresa_codigo":empresa_codigo,
                "cencos_doc":cencos_doc,
                # "codigo_empresa":codigo_empresa,
                "consecutivo":consecutivo,
                "cod_doc_asunto":cod_doc_asunto
            }
            return datos_asunto
        else:
            mensaje = MIMEText("""El Codigo de factura del correo no es correcto: """ + asunto + """ \n\n"""
                            + str(email_message['Subject']) +
                            """
                            \nTenga en cuenta que debe tener uno de los siguientes formatos, cambiando el numero de remesa correspondiente:\n"""
                            +letra+"""061 01\n
                            \nCordialmente,\n\n\nSistemas MCT. """)
            mensaje['From'] = "notificaciones@mct.com.co"
            mensaje['To'] = str(emailFrom)
            mensaje['Subject'] = "ERROR " + str(email_message['Subject'])
            self.enviar.enviar_correo("notificaciones@mct.com.co", str(emailFrom), mensaje.as_string())
            return False
        if cencos_doc_asunto and cod_doc_asunto and len(cencos_doc_asunto) == 5:
            cencos_doc_asunto = dicc_reverse_cencos[cencos_doc_asunto]
            return True
        else:
            mensaje = MIMEText("""Codigo de centro de costo inexistente:\n\n"""
                            + str(email_message['Subject']) +
                            """
                            \nCordialmente,\n\n\nSistemas MCT. """)
            mensaje['From'] = "notificaciones@mct.com.co"
            mensaje['To'] = str(emailFrom)
            mensaje['Subject'] = "ERROR " + str(email_message['Subject'])
            self.enviar.enviar_correo("notificaciones@mct.com.co", str(emailFrom), mensaje.as_string())
            return False
        # return True #nota debito

    # def insertar_digitalizado(self, modul, img, cod):
    #     global counter
    #     cod_documento = None
    #     for part in email_message.walk():
    #         if part.is_multipart():
    #             continue
    #         patron1 = re.search("text/plain|application/xml|application/pdf|multipart/mixed", part.get_content_type())
    #         patron2 = re.search("text", part.get_content_maintype())
    #         if patron1 > 0 or patron2 > 0:
    #             continue
    #         file_type = part.get_content_type().split('/')[1]
    #         print (file_type)
    #         filename = part.get_filename()
    #         if not filename:
    #             filename = name_pat.findall(part.get('Content-Type'))[0][6:-1]
    #         if filename.endswith('.pdf'):
    #             continue
    #         modulo = modul
    #         cod_documento = cod
    #         imagdoc_codigo = img
    #         # modulo = 120
    #         # cod_documento = 132930
    #         # imagdoc_codigo = 122
    #         usuario_codigo = 1
    #         # Validacion de existencia de archivo .zip
    #         sqlDigitalizado = "SELECT digitalizado_documento, digitalizado_archivo FROM tb_digitalizado WHERE modulo_codigo = "+str(modulo)+" AND  imadoc_codigo = "+str(imagdoc_codigo)+" AND digitalizado_documento = "+str(cod_documento)
    #         digitalizado_documento = ""
    #         digitalizado_archivo = ""
    #         existeDigitalizado = False
    #         try:
    #             print(sqlDigitalizado)
    #             self.cursor.execute(sqlDigitalizado)
    #             for row in  self.cursor.fetchall():
    #                 r = reg(self.cursor, row)
    #                 digitalizado_documento = r.digitalizado_documento
    #                 digitalizado_archivo = r.digitalizado_archivo
    #             existeDigitalizado = (".zip" in digitalizado_archivo)
    #             if(existeDigitalizado):
    #                 msg = "El documento: " +str(digitalizado_documento)+ ", ya cuenta con archivo .zip\n"
    #                 print(msg)
    #                 print("Archivo zip: " + str(filename))
    #                 self.enviar.correoError(cod_documento, filename, msg, modulo, emailFrom, email_message["Subject"])
    #                 break
    #         except psycopg2.Error as e:
    #             print (e.pgerror)
    #             return False
    #         outfile = 'Min_' + str(filename)
    #         print("\033[101mdigitalizado_documento: "+str(cod_documento)+"\033[0m\n")
    #         print("\033[101mArchivo zip: " + str(filename) + "\033[0m")
    #         # download the attachments from email to the designated directory
    #         # att_path = os.path.join(detach_dir, filename)
    #         # minfile = os.path.join(detach_dir, outfile)
    #         ArchivoConda = "/home/orlin/Descargas/"
    #         att_path = os.path.join(ArchivoConda, filename)
    #         minfile = os.path.join(ArchivoConda, outfile)
    #         digitalizado_ruta = os.path.join(detach_dir1, filename)
    #         if not os.path.exists(os.path.dirname(att_path)):
    #             try:
    #                 os.makedirs(os.path.dirname(att_path))
    #             except OSError as exc:
    #                 print ("error")
    #         if not os.path.isfile(att_path):
    #             # Almacenando el archivo
    #             fp = open(att_path, 'wb')
    #             temp = part.get_payload(decode=True)
    #             fp.write(temp)
    #             fp.close()
    #         sizeBytes = os.stat(att_path).st_size
    #         Msg = "Factura Almacenada en: " + str(att_path)
    #         print (Msg)
    #         Msg = "Peso en Disco Final: " + str(sizeBytes / 1024) + ' KB'
    #         print (Msg)
    #         print(detach_dir1)
    #         InsertFactura = (
    #                 "INSERT INTO tb_digitalizado "
    #                 "(modulo_codigo,digitalizado_documento,imadoc_codigo,digitalizado_archivo,digitalizado_ruta,digitalizado_fechacreacion,"
    #                 "digitalizado_fechamodificacion,usuario_codigo,digitalizado_observaciones,digitalizado_medio) "
    #                 "VALUES "
    #                 "(" + str(modulo) + ",'" + str(cod_documento) + "'," + str(imagdoc_codigo) + ",'" + filename + "','" + str(
    #             self.dicc["objeto"].get_datos_nota_ajuste(empresa_codigo, cod_doc_asunto)["digitalizado_ruta"]) + "','" + str(datenow) + "','" + str(datenow) + "'," + str(
    #             usuario_codigo) + ",'','CORREO'  ) "
    #         )
    #         print InsertFactura
    #         try:
    #             self.cursor.execute(InsertFactura)
    #             self.conexion.commit()
    #             Msg = "Factura cargada a base de datos tb_digitalizado\n._._._._._._._._._._._._._"
    #             print (Msg)
    #         except psycopg2.Error as e:
    #             Msg = (e.pgerror)
    #             print (Msg)
    #             print("ERROR FATAL EN LA CONSULTA " + str(InsertFactura))
    #             errorFile = 1
    #         # se aumenta contador para alimentar semilla del nombre de archivo
    #         counter = counter + 1
    #         if errorFile == 1:
    #             self.enviar.correoError(cod_documento, filename, Msg, modulo, emailFrom, email_message["Subject"])
    #         else:
    #             self.enviar.correoConfirmacion(cod_documento, modulo, emailFrom, email_message["Subject"])
    #             print filename
    #         datos_insert = {
    #             "filename":filename,
    #             "cod_documento":cod_documento
    #         }
    #     return datos_insert

    # def eventos(self):
    #     msg_factura = "Se carga la factura: " + str(cod_documento) + " a la plataforma Switrans"
    #     InsertFacturaEvento = (
    #             "INSERT INTO tb_eventodocumento "
    #             "(evedoc_documento,evedoc_relacion,evedoc_evento,modulo_codigo,evedoc_fechacreacion,evedoc_horacreacion,"
    #             "usuario_codigo) "
    #             "VALUES "
    #             "("+ "'" + str(cod_documento) + "','" + str(cod_documento) + "','" + str(msg_factura) + "'," + str(modulo) + ",'" + str(datenow) + "','" + str(
    #         hournow) + "'," + str(usuario_codigo) + ")")
    #     try:
    #         self.cursor.execute(InsertFacturaEvento)
    #         self.conexion.commit()
    #         Msg = "Evento cargado a base de datos tb_eventodocumento\n._._._._._._._._._._._._._"
    #         print (Msg)
    #     except psycopg2.Error as e:
    #         Msg = e.pgerror
    #         print (Msg)
    #         print("ERROR FATAL EN LA CONSULTA " + str(InsertFacturaEvento))
    #         errorFile = 1

    def analizar_correo(self, correo, password, carpeta):
        global counteraux,email_message, name_pat, cambio
        cambio = re.sub("_"," ",carpeta)
        M = imaplib.IMAP4_SSL(host, clave)
        M.login(correo, password)
        M.select(carpeta)
        name_pat = re.compile('name=".*"')
        Msg = "\033[92mListo para traer mensajes sin leer de la cuenta "+correo+" de la carpeta " +cambio.upper()+"\033[0m"
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
                # eMail = re.findall(r'\W*(fe)\W*@[\w\.-]+', email_message["From"])
                # print("este es email de "+carpeta+":" + email_message['From'])
                # if(len(eMail) <= 0):
                #     Msg = "El Correo es enviado desde un contacto NO valido debe ser fe@mct.com.co"
                #     print (Msg)
                #     break
                Subj = decode_header(email_message['Subject'])[0][0]
                # if self.validar_asunto_correo(Subj,"E") is False:
                if 1>2:
                    Msg = "El Asunto no es correcto o esta mal formado (no tiene el numero de documento soporte), se envía correo y se detiene el proceso"
                    print(Msg)
                    break
                else:
                    # print ("Cencos: " + str(cencos_doc_asunto))
                    # print ("reso: " + str(cod_doc_asunto))
                    # if self.dicc["objeto"].getDatos(cencos_doc_asunto, cod_doc_asunto) is False:
                    if 1>2:
                        Msg = "No se lograron obtener datos de la " + cambio.upper()
                        print (Msg)
                        break
                    print "ejecuto el insert_digitalizado"
                    # self.insertar_digitalizado(any, any, any)
                    # digitalizado
            counteraux = counteraux + 1

        if counter == 0 :
            Msg = "\033[91mNo existen mensajes sin leer en la cuenta "+ correo + " en la carpeta "+ cambio.upper()+"\033[0m"
            print (Msg)
        M.close()
        M.logout()
        return Subj
        Msg = "\033[93m*********************************************** Proceso Finalizado " +cambio.upper()+ " ***********************************************\033[0m\n\n"
        print (Msg)
