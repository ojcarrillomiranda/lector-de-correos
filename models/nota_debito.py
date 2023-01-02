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
ano = 0
mes = 0
dia = 0
Msg = None

empresa_codigo_general_nota_debito = {
  "01": "1",
  "13": "12",
  "02": "2"
}


parser = HTMLParser.HTMLParser()
datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

config = ConfigParser()
config.read('config/config.ini')
correo = config.get('conf','CORREO_O')
password =  config.get('conf','PASS')
carpeta = config.get('conf','IMAP_ND')
cargue_string = config.get('conf','CARGUE_NOTA_DEBITO')
error_string = config.get('conf','ERROR_NOTA_DEBITO')

class NotaDebito:
    def leer_correos_nota_debito(self, con, env):
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
                valida_asunto = Funcion(cursor, conexion, enviar).validar_asunto_correo("ND", asunto)
                empresa_codigo_contable = valida_asunto["empresa_codigo_contable"]
                consecutivo = valida_asunto["consecutivo"]
                empresa = valida_asunto["empresa"]
                codigo_empresa = empresa_codigo_general_nota_debito[empresa_codigo_contable]
                if valida_asunto is False:
                    mensaje = "\033[91mEl Asunto no es correcto o esta mal formado (no tiene el numero de documento soporte), se envía correo y se detiene el proceso\033[0m"
                    print mensaje
                    break
                datos_nota_debito = self.get_datos_nota_debito(cursor, codigo_empresa, consecutivo, empresa)
                nota_debito = datos_nota_debito["codigo_nota_debito"]
                ruta = datos_nota_debito["digitalizado_ruta"]
                if datos_nota_debito is False:
                    mensaje = "\033[91No se lograron obtener datos de la " + cambio.upper() + "\033[0m"
                    print mensaje
                    break
                insert_digi = Funcion(cursor, conexion, enviar).insertar_digitalizado(256, 140, nota_debito, ruta, filename, cargue_string, error_string)
                if insert_digi:
                    archivo_zip = insert_digi
                    self.actuLizar_nota_debito(cursor, conexion, archivo_zip, nota_debito)
                    Funcion(cursor, conexion, enviar).eventos(nota_debito, 256)
            mensaje = "\033[91mNo existen mensajes sin leer en la cuenta " + correo + " en la carpeta " + cambio.upper() + "\033[0m"
            print (mensaje)
            mensaje = "\033[93m*********************************************** Proceso Finalizado " + cambio.upper() + " ***********************************************\033[0m\n\n"
            print (mensaje)




    def get_datos_nota_debito(self, cursor, empresa_codigo, consecutivo, empresa):
        global ruta_servidor, digitalizado_ruta, codigo_nota_debito, cencos_nota_debito, fecha_nota_debito, ano, mes, dia
        nota_debito_sql = "SELECT notadebito_codigo, nd.cencos_codigo, nd.notadebito_fechacreacion FROM tb_notadebito nd LEFT JOIN tb_centrocosto cc ON nd.cencos_codigo = cc.cencos_codigo WHERE cc.empresa_codigo = " + str(empresa_codigo) + " AND nd.notadebito_cencoscodigo = " + str(consecutivo)
        mensaje = "\033[94mConsulta de la nota debito para la empresa " + str(empresa_codigo) + " ccNotaDebito " + str(consecutivo) + " ejecutada con exito\033[0m"
        try:
            print "\033[94m######################### SQL get_datos() #################################\033[0m"
            print ("\033[94m" + nota_debito_sql) + "\033[0m"
            cursor.execute(nota_debito_sql)
            print (mensaje)
            outfile = '.'
            ContAux = 0
            for row in cursor.fetchall():
                r = reg(cursor, row)
                codigo_nota_debito = r.notadebito_codigo
                cencos_nota_debito = r.cencos_codigo
                fecha_nota_debito = str(r.notadebito_fechacreacion)
                ano = fecha_nota_debito.split("-")[0]
                mes = fecha_nota_debito.split("-")[1]
                dia = fecha_nota_debito.split("-")[2]
                digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/facturas/' + str(empresa) + '/' + str(ano) + '/' + str(
                    cencos_nota_debito) + '/' + str(mes) + '/'
                ContAux = ContAux + 1
                print "\033[92m######################### RESULTADO SQL get_datos() #################################\033[0m\n" + "\033[92mnotadebito_codigo = " + str(
                    codigo_nota_debito) + "\n" + "cencos_codigo = " + str(cencos_nota_debito) + "\n" + "notadebito_fechacreacion = " + fecha_nota_debito + "\033[0m\n"
            if ContAux > 0:
                datos_nota_credito = {
                    "codigo_nota_debito": codigo_nota_debito,
                    "digitalizado_ruta": digitalizado_ruta
                }
                return datos_nota_credito
            else:
                return False
        except psycopg2.Error as e:
            print(e.pgerror)
            return False

    def actuLizar_nota_debito(self, cursor, conexion, filename, nota_debito):
        update_nota_debito = (
                "UPDATE tb_notacontabilidad SET archivo_facturacionelectronica = '" + filename + "' WHERE notcon_codigo = " + str(
            nota_debito)
        )
        try:
            cursor.execute(update_nota_debito)
            conexion.commit()
            mensaje = "\033[93mNota debito Actualizada a base de datos tb_notacontabilidad, campo: archivo_facturacionelectronica \n._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._\033[0m"
            print(mensaje)
        except psycopg2.Error as e:
            mensaje = e.pgerror
            print(mensaje)
            print("ERROR FATAL EN LA CONSULTA " + str(update_nota_debito))
            errorFile = 1
