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
from models.reg import reg
from common.funciones import Funcion

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
# empresaCodCont = {
#   "MCT": "01",
#   "FERRICAR": "13",
#   "MARKETING": "02"
# }
# empresaCodGeneral = {
#   "MCT": "1",
#   "FERRICAR": "12",
#   "MARKETING": "2"
# }

parser = HTMLParser.HTMLParser()
datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)


config = ConfigParser()
config.read('config/config.ini')
correo = config.get('conf','CORREO_O')
password =  config.get('conf','PASS')
carpeta = config.get('conf','IMAP_FAC')

class Factura:

    def leer_correos_factura(self, con, env):
        cursor = con.cursor()
        conexion = con
        enviar = env

        dicReverseCC = Funcion(cursor,conexion,enviar).get_datos_centro_costo()

        analisa_correo = Funcion(cursor, conexion, enviar).analizar_correo(correo, password, carpeta)

        asunto = Funcion(cursor, conexion, enviar).validar_asunto_correo(analisa_correo, "E")
        cencos_codigo = asunto["cencos_doc"]
        documento = asunto["cod_doc_asunto"]

        self.get_datos_factura(cursor, cencos_codigo, documento)

        # validar_correo = Funcion(cursor,conexion,enviar).validar_asunto_correo(analisa_correo,"E")
        # cencos_doc_asunto = validar_correo["cencos_doc_asunto"]
        # cod_doc_asunto = validar_correo["cod_doc_asunto"]
        #
        # # dato = self.get_datos_factura(26, 3825)["codigo_factura"]
        # dato = self.get_datos_factura(cursor, cencos_doc_asunto, cod_doc_asunto)["codigo_factura"]
        # print "este es el dato ",dato
        # insert_digi = Funcion(cursor, conexion, enviar).insertar_digitalizado(120, 122, dato)
        #
        # update = self.update_factura(insert_digi["filename"],insert_digi["cod_documento"])

    def get_datos_factura(self,cursor,cencos_codigo, documento):
        global detach_dir, detach_dir1, codigo_factura, cencos_factura, fecha_factura, ano, mes, dia, facturaSubj
        FacturaSQL = "Select factura_codigo, cencos_codigo, factura_fechacreacion from tb_factura where cencos_codigo = '" + str(cencos_codigo) + "' AND factura_numerodocumento = '" + str(documento) + "'"
        Msg = "Consulta de la Factura CC " + str(cencos_codigo) + " DigCCResolucion " + str(documento) + " ejecutada con éxito"
        try:
            print (FacturaSQL)
            cursor.execute(FacturaSQL)
            print (Msg)
            outfile = '.'
            ContAux = 0
            for row in cursor.fetchall():
                print ("paso a consultar")
                print ("##########################")
                r = reg(cursor, row)
                codigo_factura = r.factura_codigo
                cencos_factura = r.cencos_codigo
                fecha_factura = str(r.factura_fechacreacion)
                ano = fecha_factura.split("-")[0]
                mes = fecha_factura.split("-")[1]
                dia = fecha_factura.split("-")[2]
                ruta_servidor = '/documents/facturas/' + str(Emp) + '/' + str(ano) + '/' + str(cencos_factura) + '/' + str(mes) + '/'
                digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/facturas/' + str(Emp) + '/' + str(ano) + '/' + str(cencos_factura) + '/' + str(mes) + '/'
                ContAux = ContAux + 1
            if ContAux > 0:
                datos_factura = {
                    "codigo_factura":codigo_factura,
                    "cencos_factura":cencos_factura,
                    "fecha_factura":fecha_factura,
                    "ruta_servidor":ruta_servidor,
                    "digitalizado_ruta":digitalizado_ruta
                }
                print "estos son los datos de la factura->",datos_factura
                return datos_factura
            else:
                return False
        except psycopg2.Error as e:
            print (e.pgerror)
            return False

    # def update_factura(self, filename, cod_documento):
    #     update_factura = (
    #             "UPDATE tb_factura SET archivo_facturacionelectronica = '" + filename + "' WHERE factura_codigo = " + str(
    #         cod_documento)
    #     )
    #     try:
    #         cursor.execute(update_factura)
    #         conexion.commit()
    #         Msg = "Factura Actualizada a base de datos tb_factura, campo: archivo_facturacionelectronica \n._._._._._._._._._._._._._"
    #         print (Msg)
    #     except psycopg2.Error as e:
    #         Msg = (e.pgerror)
    #         print (Msg)
    #         print("ERROR FATAL EN LA CONSULTA " + str(update_factura))
    #         errorFile = 1





