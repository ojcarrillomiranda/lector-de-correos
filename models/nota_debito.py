#! /usr/bin/env python
# -*- coding: utf-8 -*-
from configparser import ConfigParser
import imaplib, email, base64
from email.header import decode_header
import smtplib
from common.funciones import Funcion
from conexiones.EnviarCorreo import EnviarCorreo
import HTMLParser
import os, sys
import psycopg2
import psycopg2.extras
import psycopg2.extensions

from conexiones.ConexionDB import Conexion
from models.reg import reg
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
# ruta_servidor = '.'
# digitalizado_ruta = '.'
dicReverseCC = {}
filename = None
outfile = None
ano = 0
mes = 0
dia = 0
Msg = None
codigo_nota_debito = None
cencos_nota_debito = None
fecha_nota_debito = None
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
        # def __init__(self, conexion):
        #     self.cursor = conexion.cursor()
    def leer_correos_nota_debito(self, conexion):
        cursor = conexion.cursor()
        Funcion().setReverseCC(cursor)
        def get_datos_nota_debito(empCodigo, consecutivoCC):
            global ruta_servidor, digitalizado_ruta, codigo_nota_debito, cencos_nota_debito, fecha_nota_debito, ano, mes, dia
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
                    print ("##########################")
                    r = reg(cursor, row)
                    codigo_nota_debito = r.notadebito_codigo
                    cencos_nota_debito = r.cencos_codigo
                    fecha_nota_debito = str(r.notadebito_fechacreacion)
                    ano = fecha_nota_debito.split("-")[0]
                    mes = fecha_nota_debito.split("-")[1]
                    dia = fecha_nota_debito.split("-")[2]
                    ruta_servidor = '/documents/facturas/' + str(Emp) + '/' + str(ano) + '/' + str(cencos_nota_debito) + '/' + str(mes) + '/'
                    digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/facturas/' + str(Emp) + '/' + str(ano) + '/' + str(
                        cencos_nota_debito) + '/' + str(mes) + '/'
                    ContAux = ContAux + 1        
                if ContAux > 0:
                    datos_nota_credito = {
                        "codigo_nota_debito": codigo_nota_debito,
                        "cencos_nota_debito": cencos_nota_debito,
                        "fecha_nota_debito": fecha_nota_debito,
                        "ruta_servidor": ruta_servidor,
                        "digitalizado_ruta": digitalizado_ruta
                    }
                    print datos_nota_credito,"\n"
                    return datos_nota_credito
                else:
                    return False
            except psycopg2.Error as e:
                print(e.pgerror)
                return False
        get_datos_nota_debito(1,31)

        def update_nota_debito():
            updateNotaDebito = (
                    "UPDATE tb_notacontabilidad SET archivo_facturacionelectronica = '" + filename + "' WHERE notcon_codigo = " + str(
                cod_documento)
            )
            try:
                cursor.execute(updateNotaDebito)
                conn.commit()
                Msg = "Nota debito Actualizada a base de datos tb_notacontabilidad, campo: archivo_facturacionelectronica \n._._._._._._._._._._._._._"
                logfile.write(Msg + '\n\n')
                print(Msg)
                # except Exception, e:
            except psycopg2.Error as e:
                Msg = e.pgerror
                logfile.write(Msg)
                print(Msg)
                print("ERROR FATAL EN LA CONSULTA " + str(updateNotaDebito))
                errorFile = 1

        # dicReverseCC = Funcion().setReverseCC()
        # print("dicReverseCC nota credito: "+str(dicReverseCC))
        # Funcion().analizar_correo(config.get('conf','CORREO_O'), config.get('conf','PASS'), config.get('conf','IMAP_ND'))