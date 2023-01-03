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
ano = 0
mes = 0
dia = 0
Msg = None

parser = HTMLParser.HTMLParser()
datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

config = ConfigParser()
config.read('config/config.ini')
correo = config.get('conf','CORREO_O')
password =  config.get('conf','PASS')
carpeta = config.get('conf','IMAP_FAC')
cargue_string = config.get('correo_confirmacion','CARGUE_FACTURA')
error_string = config.get('correo_error','ERROR_FACTURA')


class Factura:
    def leer_correos_factura(self, con, env):
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
                valida_asunto = Funcion(cursor, conexion, enviar).validar_asunto_correo("E", asunto)
                cencos_codigo = valida_asunto["cencos_doc_asunto"]
                documento =  valida_asunto["cod_doc_asunto"]
                empresa = valida_asunto["empresa"]
                if valida_asunto is False:
                    mensaje = "\033[91mEl Asunto no es correcto o esta mal formado (no tiene el numero de documento soporte), se envía correo y se detiene el proceso\033[0m"
                    print mensaje
                    break
                datos_factura = self.get_datos_factura(cursor, cencos_codigo, documento, empresa)
                factura = datos_factura["codigo_factura"]
                ruta = datos_factura["digitalizado_ruta"]
                if datos_factura is False:
                    mensaje = "\033[91No se lograron obtener datos de la " + cambio.upper() + "\033[0m"
                    print mensaje
                    break
                insert_digi = Funcion(cursor, conexion, enviar).insertar_digitalizado(120, 122, factura, ruta, filename, cargue_string, error_string)
                if insert_digi:
                    archivo_zip = insert_digi
                    self.actualizar_factura(cursor, conexion, archivo_zip, factura)
                    Funcion(cursor, conexion, enviar).eventos(factura, 120)
            mensaje = "\033[91mNo existen mensajes sin leer en la cuenta " + correo + " en la carpeta " + cambio.upper() + "\033[0m"
            print (mensaje)
            mensaje = "\033[93m*********************************************** Proceso Finalizado " + cambio.upper() + " ***********************************************\033[0m\n\n"
            print (mensaje)

    def get_datos_factura(self,cursor,cencos_codigo, documento, empresa):
        global detach_dir, detach_dir1, codigo_factura, cencos_factura, fecha_factura, ano, mes, dia, facturaSubj
        facturas_sql = "Select factura_codigo, cencos_codigo, factura_fechacreacion from tb_factura where cencos_codigo = '" + str(cencos_codigo) + "' AND factura_numerodocumento = '" + str(documento) + "'"
        mensaje = "\033[94mConsulta de la Factura CC " + str(cencos_codigo) + " DigCCResolucion " + str(documento) + " ejecutada con éxito\033[0m"
        try:
            print "\033[94m######################### SQL get_datos() #################################\033[0m"
            print ("\033[94m" + facturas_sql + "\033[0m")
            cursor.execute(facturas_sql)
            print (mensaje)
            outfile = '.'
            ContAux = 0
            for row in cursor.fetchall():
                r = reg(cursor, row)
                codigo_factura = r.factura_codigo
                cencos_factura = r.cencos_codigo
                fecha_factura = str(r.factura_fechacreacion)
                ano = fecha_factura.split("-")[0]
                mes = fecha_factura.split("-")[1]
                dia = fecha_factura.split("-")[2]
                digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/facturas/' + str(empresa) + '/' + str(ano) + '/' + str(cencos_factura) + '/' + str(mes) + '/'
                ContAux = ContAux + 1
                print "\033[92m######################### RESULTADO SQL get_datos() #################################\033[0m\n" + "\033[92mcodigo_factura = " + str(codigo_factura) + "\n" + "cencos_codigo = " + str(cencos_factura) +"\n"+ "factura_fechacreacion = " + fecha_factura + "\033[0m\n"
            if ContAux > 0:
                datos_factura = {
                    "codigo_factura":codigo_factura,
                    "digitalizado_ruta":digitalizado_ruta
                }
                return datos_factura
            else:
                return False
        except psycopg2.Error as e:
            print (e.pgerror)
            return False

    def actualizar_factura(self, cursor, conexion, filename, factura):
        update_factura = (
                "UPDATE tb_factura SET archivo_facturacionelectronica = '" + filename + "' WHERE factura_codigo = " + str(
            factura)
        )
        try:
            cursor.execute(update_factura)
            conexion.commit()
            mensaje = "\033[93mFactura Actualizada a base de datos tb_factura, campo: archivo_facturacionelectronica \n._._._._._._._._._._._._._ \033[0m"
            print (mensaje)
        except psycopg2.Error as e:
            mensaje = (e.pgerror)
            print (mensaje)
            print("ERROR FATAL EN LA CONSULTA " + str(update_factura))
            errorFile = 1





