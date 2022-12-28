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

# fecha actual
datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)


Emp = ""
# emailFrom = "doc_soporte@mct.com.co"
emailFrom = "orlin.carrillo@mct.com.co"
counter = 0
counteraux = 1
imgType = 0
usuario_codigo = 0
ruta_servidor = '.'
eMailAut = []
dicReverseCC = {}
remesaSubj = None
ccRemesaSubj = None
codCcRemesaSubj = None
filename = None
outfile = None
ano = 0
mes = 0
dia = 0
codigo_remesa = None
cencos_remesa = 0
fecha_remesa = None
Msg = None

codigo_factura = None
cencos_factura = None
fecha_factura = None
facturaSubj = None
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

class DocumentoSoporte(object):
        # def __init__(self, conexion):
        #     self.cursor = conexion.cursor()
    def leer_correos_documento_soporte(self, conexion):
        cursor = conexion.cursor()
        Funcion().setReverseCC(cursor)
        def get_datos_doc_soporte(empCodigo, codCcDocumentoSoporte, ccDocumentoSubj):
            print "empresa codigo ",empCodigo
            print "cencos codigo ",codCcDocumentoSoporte
            print "item resolucion ",ccDocumentoSubj
            global ruta_servidor, digitalizado_ruta, codigo_docsop, cencos_docsop, fecha_docsop, ano, mes, dia, docsopSubj
            docsopSQL = "SELECT ds.docsop_codigo as docsop_codigo, ds.cencos_codigo, ds.docsop_fechacreacion AS docsop_fechacreacion FROM tb_documentosoporte ds INNER JOIN tb_itemresoluciondocumentosoporte irds ON irds.itresoldocsop_codigo = ds.itresoldocsop_codigo WHERE ds.empresa_codigo = " + str(empCodigo) + " AND ds.cencos_codigo = " + str(ccDocumentoSubj) + " AND irds.itresoldocsop_numero = " + str(codCcDocumentoSoporte)
            Msg = "Consulta del Documento Soporte Empresa " + str(empCodigo) + " DigCCResolucion " + str(codCcDocumentoSoporte) + " ejecutada con éxito"
            try:
                print (docsopSQL)
                cursor.execute(docsopSQL)
                print (Msg)
                outfile = '.'
                ContAux = 0
                for row in cursor.fetchall():
                    print ("paso a consultar")
                    r = reg(cursor, row)
                    docsopSubj = r.docsop_codigo
                    codigo_docsop = r.docsop_codigo
                    cencos_docsop = r.cencos_codigo
                    fecha_docsop = str(r.docsop_fechacreacion)
                    ano = fecha_docsop.split("-")[0]
                    mes = fecha_docsop.split("-")[1]
                    dia = fecha_docsop.split("-")[2]
                    print ("##########################")
                    print (codigo_docsop)
                    ruta_servidor = '/documents/documentosoporte/' + str(Emp) + '/' + str(ano) + '/' + str(cencos_docsop) + '/' + str(
                        mes) + '/'
                    digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/documentosoporte/' + str(Emp) + '/' + str(
                        ano) + '/' + str(cencos_docsop) + '/' + str(mes) + '/'
                    ContAux = ContAux + 1
                if ContAux > 0:
                    datos_doc_soporte = {
                        "codigo_docsop": codigo_docsop,
                        "cencos_docsop": cencos_docsop,
                        "fecha_docsop": fecha_docsop,
                        "ruta_servidor": ruta_servidor,
                        "digitalizado_ruta": digitalizado_ruta
                    }
                    print datos_doc_soporte,"\n"
                    return datos_doc_soporte
                else:
                    return False
            except psycopg2.Error as e:
                print(e.pgerror)
            return False
        get_datos_doc_soporte(12, 39, 297)

        def update_doc_soporte():
            updateDocSoporte = (
                    "UPDATE tb_documentosoporte SET archivo_documentosoportelectronico = '" + filename + "' WHERE docsop_codigo = " + str(
                cod_documento))
            try:
                cursor.execute(updateDocSoporte)
                conn.commit()
                Msg = "Documento Soporte Actualizada a base de datos tb_documentosoporte, campo: archivo_documentosoportelectronico \n._._._._._._._._._._._._._"
                logfile.write(Msg + '\n\n')
                print
                Msg
            except psycopg2.Error as e:
                Msg = e.pgerror
                logfile.write(Msg)
                print Msg
                print("ERROR FATAL EN LA CONSULTA " + str(updateDocSoporte))
                errorFile = 1

