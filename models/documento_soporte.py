#! /usr/bin/env python
# -*- coding: utf-8 -*-
from configparser import ConfigParser
import imaplib, email
from email.header import decode_header
import HTMLParser
import os
import psycopg2
import psycopg2.extras
import psycopg2.extensions

from common.funciones import Funcion
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)
import datetime, re
import email
from email.MIMEText import MIMEText
from models.reg import reg

datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

Emp = ""
ano = 0
mes = 0
dia = 0

parser = HTMLParser.HTMLParser()
datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

config = ConfigParser()
config.read('config/config.ini')
correo = config.get('conf','CORREO_O')
password =  config.get('conf','PASS')
carpeta = config.get('conf','IMAP_DS')
cargue_string = config.get('correo_confirmacion','CARGUE_DOC_SOPORTE')
error_string = config.get('correo_error','ERROR_DOC_SOPORTE')

class DocumentoSoporte:

    def leer_correos_documento_soporte(self, con, env):
        cursor = con.cursor()
        conexion = con
        enviar = env
        cambio = re.sub("_", " ", carpeta)

        Funcion(cursor, conexion, enviar).get_datos_centro_costo()
        data_correos = Funcion(cursor, conexion, enviar).analizar_correo(correo, password, carpeta)

        if data_correos != None:
            for dato in data_correos:
                asunto = dato["asunto"]
                filename = dato["filename"]

                valida_asunto = Funcion(cursor, conexion, enviar).validar_asunto_correo("D", asunto)
                empresa_codigo = valida_asunto["empresa_codigo"]
                documento_soporte = valida_asunto["cod_doc_asunto"]
                cencos_documento = valida_asunto["cencos_doc_asunto"]
                empresa = valida_asunto["empresa"]
                if valida_asunto is False:
                    mensaje = "\033[91mEl Asunto no es correcto o esta mal formado (no tiene el numero de documento soporte), se envía correo y se detiene el proceso\033[0m"
                    print mensaje
                    break
                datos_doc_soporte = self.get_datos_doc_soporte(cursor, empresa_codigo, documento_soporte, cencos_documento, empresa)
                doc_soporte = datos_doc_soporte["codigo_docsop"]
                ruta = datos_doc_soporte["digitalizado_ruta"]
                if datos_doc_soporte is False:
                    mensaje = "\033[91No se lograron obtener datos de la " + cambio.upper() + "\033[0m"
                    print mensaje
                    break
                insert_digi = Funcion(cursor, conexion, enviar).insertar_digitalizado(299, 141, doc_soporte, ruta, filename, cargue_string, error_string)
                if insert_digi:
                    archivo_zip = insert_digi
                    self.actualizar_doc_soporte(cursor, conexion, archivo_zip, doc_soporte)
                    Funcion(cursor, conexion, enviar).eventos(doc_soporte, 299)
            mensaje = "\033[91mNo existen mensajes sin leer en la cuenta " + correo + " en la carpeta " + cambio.upper() + "\033[0m"
            print (mensaje)
            mensaje = "\033[93m*********************************************** Proceso Finalizado " + cambio.upper() + " ***********************************************\033[0m\n\n"
            print (mensaje)

    def get_datos_doc_soporte(self, cursor, empresa_codigo, documento_soporte, cencos_documento, empresa):
        global ruta_servidor, digitalizado_ruta, codigo_docsop, cencos_docsop, fecha_docsop, ano, mes, dia, docsopSubj
        doc_soporte_sql = "SELECT ds.docsop_codigo as docsop_codigo, ds.cencos_codigo, ds.docsop_fechacreacion AS docsop_fechacreacion FROM tb_documentosoporte ds INNER JOIN tb_itemresoluciondocumentosoporte irds ON irds.itresoldocsop_codigo = ds.itresoldocsop_codigo WHERE ds.empresa_codigo = " + str(empresa_codigo) + " AND ds.cencos_codigo = " + str(cencos_documento) + " AND irds.itresoldocsop_numero = " + str(documento_soporte)
        mensaje = "\033[94mConsulta del Documento Soporte Empresa " + str(empresa_codigo) + " DigCCResolucion " + str(documento_soporte) + " ejecutada con éxito\033[0m"
        try:
            print "\033[94m######################### SQL get_datos() #################################\033[0m"
            print ("\033[94m" + doc_soporte_sql + "\033[0m")
            cursor.execute(doc_soporte_sql)
            print (mensaje)
            outfile = '.'
            ContAux = 0
            for row in cursor.fetchall():
                r = reg(cursor, row)
                codigo_docsop = r.docsop_codigo
                cencos_docsop = r.cencos_codigo
                fecha_docsop = str(r.docsop_fechacreacion)
                ano = fecha_docsop.split("-")[0]
                mes = fecha_docsop.split("-")[1]
                dia = fecha_docsop.split("-")[2]
                digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/documentosoporte/' + str(empresa) + '/' + str(
                    ano) + '/' + str(cencos_docsop) + '/' + str(mes) + '/'
                ContAux = ContAux + 1
                print "\033[92m######################### RESULTADO SQL get_datos() #################################\033[0m\n" + "\033[92mdocsop_codigo = " + str(
                    codigo_docsop) + "\n" + "cencos_codigo = " + str(cencos_docsop) + "\n" + "docsop_fechacreacion = " + fecha_docsop + "\033[0m\n"
            if ContAux > 0:
                datos_doc_soporte = {
                    "codigo_docsop": codigo_docsop,
                    "digitalizado_ruta": digitalizado_ruta
                }
                return datos_doc_soporte
            else:
                return False
        except psycopg2.Error as e:
            print(e.pgerror)
        return False

    def actualizar_doc_soporte(self, cursor, conexion, filename, doc_soporte):
        update_doc_soporte = (
                "UPDATE tb_documentosoporte SET archivo_documentosoportelectronico = '" + filename + "' WHERE docsop_codigo = " + str(
            doc_soporte))
        try:
            cursor.execute(update_doc_soporte)
            conexion.commit()
            mensaje = "\033[93mDocumento Soporte Actualizada a base de datos tb_documentosoporte, campo: archivo_documentosoportelectronico \n._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._\033[0m"
            print mensaje
        except psycopg2.Error as e:
            mensaje = e.pgerror
            print mensaje
            print("ERROR FATAL EN LA CONSULTA " + str(update_doc_soporte))
            errorFile = 1

