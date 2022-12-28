#! /usr/bin/env python
# -*- coding: utf-8 -*-
import imaplib, email
from email.header import decode_header
import HTMLParser
import os
import psycopg2
import psycopg2.extras
import psycopg2.extensions
from common.funciones import Funcion
from models.reg import reg
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
        # def __init__(self, conexion):
        #     self.cursor = conexion.cursor()
    def leer_correos_nota_ajuste(self, conexion):
        cursor = conexion.cursor()
        Funcion().setReverseCC(cursor)
        def get_datos_nota_ajuste(empCodigo, codCcNotaAjuste):
            print (empCodigo)
            print (codCcNotaAjuste)
            global detach_dir, detach_dir1, cod_nota_ajuste, cencos_nota_ajuste, fecha_nota_ajuste, ano, mes, dia
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
                    cod_nota_ajuste = r.notaju_codigo
                    cencos_nota_ajuste = r.cencos_codigo
                    fecha_nota_ajuste = str(r.notaju_fechacreacion)
                    ano = fecha_nota_ajuste.split("-")[0]
                    mes = fecha_nota_ajuste.split("-")[1]
                    dia = fecha_nota_ajuste.split("-")[2]
                    print ("##########################")
                    print( cod_nota_ajuste)
                    ruta_servidor = '/documents/documentosoporte/' + str(Emp) + '/' + str(ano) + '/' + str(cencos_nota_ajuste) + '/' + str(
                        mes) + '/'
                    digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/documentosoporte/' + str(Emp) + '/' + str(
                        ano) + '/' + str(cencos_nota_ajuste) + '/' + str(mes) + '/'
                ContAux = ContAux + 1
                if ContAux > 0:
                    datos_nota_ajuste = {
                        "cod_nota_ajuste": cod_nota_ajuste,
                        "cencos_nota_ajuste": cencos_nota_ajuste,
                        "fecha_nota_ajuste": fecha_nota_ajuste,
                        "ruta_servidor": ruta_servidor,
                        "digitalizado_ruta": digitalizado_ruta
                    }
                    print datos_nota_ajuste,"\n"
                    return datos_nota_ajuste
                else:
                    return False
            except psycopg2.Error as e:
                print(e.pgerror)
            return False
        get_datos_nota_ajuste(1,18)

        def update_nota_ajuste():
            updateNotaAjuste = (
                    "UPDATE tb_notaajuste SET archivo_notaajuste_electronico = '" + filename + "' WHERE notaju_codigo = " + str(
                cod_documento))
            try:
                cursor.execute(updateNotaAjuste)
                conn.commit()
                Msg = "Nota Ajuste Actualizada a base de datos tb_notaajuste, campo: archivo_notaajuste_electronico \n._._._._._._._._._._._._._"
                logfile.write(Msg + '\n\n')
                print
                Msg
            except psycopg2.Error as e:
                Msg = e.pgerror
                logfile.write(Msg)
                print Msg
                print("ERROR FATAL EN LA CONSULTA " + str(updateNotaAjuste))
                errorFile = 1