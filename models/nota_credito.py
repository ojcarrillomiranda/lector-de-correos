#! /usr/bin/env python
# -*- coding: utf-8 -*-
from configparser import ConfigParser
import datetime, re
from configparser import ConfigParser
from models.reg import reg
from common.funciones import Funcion

Emp = ""
ano = 0
mes = 0
dia = 0

datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

config = ConfigParser()
config.read('config/config.ini')
correo = config.get('personal','CORREO_O')
password =  config.get('personal','PASS')
carpeta = config.get('carga_fe','IMAP_NC')
cargue_string = config.get('correo_confirmacion','CARGUE_NOTA_CREDITO')
error_string = config.get('correo_error','ERROR_NOTA_CREDITO')


class NotaCredito:
    def leer_correos_nota_credito(self, con, env):
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
                valida_asunto = Funcion(cursor, conexion, enviar).validar_asunto_correo("C", asunto)
                cod_doc_asunto = valida_asunto["cod_doc_asunto"]
                empresa_codigo = valida_asunto["empresa_codigo"]
                empresa = valida_asunto["empresa"]
                if valida_asunto is False:
                    mensaje = "\033[91mEl Asunto no es correcto o esta mal formado (no tiene el numero de documento soporte), se envía correo y se detiene el proceso\033[0m"
                    print(mensaje)
                    break
                datos_nota_credito = self.get_datos_nota_credito(cursor, empresa_codigo, cod_doc_asunto, empresa)
                nota_credito = datos_nota_credito["codigo_nota"]
                ruta = datos_nota_credito["digitalizado_ruta"]
                if datos_nota_credito is False:
                    mensaje = "\033[91No se lograron obtener datos de la " + cambio.upper() + "\033[0m"
                    print(mensaje)
                    break
                insert_digi = Funcion(cursor, conexion, enviar).insertar_digitalizado(129, 135, nota_credito, ruta, filename, cargue_string, error_string)
                if insert_digi:
                    archivo_zip = insert_digi
                    self.actualizar_nota_credito(cursor, conexion, archivo_zip, nota_credito)
                    Funcion(cursor, conexion, enviar).eventos(nota_credito, 129)
            mensaje = "\033[91mNo existen mensajes sin leer en la cuenta " + correo + " en la carpeta " + cambio.upper() + "\033[0m"
            print (mensaje)
            mensaje = "\033[93m*********************************************** Proceso Finalizado " + cambio.upper() + " ***********************************************\033[0m\n\n"
            print (mensaje)

    def get_datos_nota_credito(self, cursor, empresa_codigo, cencos_codigo, empresa):
        global detach_dir, detach_dir1, codigo_nota, cencos_nota, fecha_nota, ano, mes, dia, notaSubj
        nota_credito_sql = "Select notcon_codigo as nota_codigo, cencos_codigo, notcon_fechacreacion AS nota_fechacreacion from tb_notacontabilidad where empresa_codigo = '" + str(
            empresa_codigo) + "' AND notcon_cencoscodigo = '" + str(cencos_codigo) + "'"
        mensaje = "\033[94mConsulta de la nota credito Empresa " + str(empresa_codigo) + " DigCCResolucion " + str(cencos_codigo) + " ejecutada con éxito\033[0m"
        try:
            print("\033[94m######################### SQL get_datos() #################################\033[0m")
            print ("\033[94m" + nota_credito_sql) + "\033[0m"
            cursor.execute(nota_credito_sql)
            print (mensaje)
            outfile = '.'
            ContAux = 0
            for row in cursor.fetchall():
                r = reg(cursor, row)
                codigo_nota = r.nota_codigo
                cencos_nota = r.cencos_codigo
                fecha_nota_credito = str(r.nota_fechacreacion)
                ano = fecha_nota_credito.split("-")[0]
                mes = fecha_nota_credito.split("-")[1]
                dia = fecha_nota_credito.split("-")[2]
                digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/facturas/' + str(empresa) + '/' + str(
                    ano) + '/' + str(cencos_nota) + '/' + str(mes) + '/'
                ContAux = ContAux + 1
                print("\033[92m######################### RESULTADO SQL get_datos() #################################\033[0m\n" + "\033[92mnotcon_codigo = " + str(
                    codigo_nota) + "\n" + "cencos_codigo = " + str(
                    cencos_nota) + "\n" + "notcon_fechacreacion = " + fecha_nota_credito + "\033[0m\n")
            if ContAux > 0:
                datos_nota_credito = {
                    "codigo_nota": codigo_nota,
                    "digitalizado_ruta": digitalizado_ruta
                }
                return datos_nota_credito
            else:
                return False
        except psycopg2.Error as e:
            print (e.pgerror)
            return False

    def actualizar_nota_credito(self, cursor, conexion, filename, nota_credito):
        update_nota_credito = (
                    "UPDATE tb_notacontabilidad SET archivo_facturacionelectronica = '" + filename + "' WHERE notcon_codigo = " + str(
                nota_credito))
        try:
            cursor.execute(update_nota_credito)
            conexion.commit()
            mensaje = "\033[93mNota Credito Actualizada a base de datos tb_notacontabilidad, campo: archivo_facturacionelectronica \n._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._\033[0m"
            print(mensaje)
        except psycopg2.Error as e:
            mensaje = e.pgerror
            print(mensaje)
            print("ERROR FATAL EN LA CONSULTA " + str(update_nota_credito))
            errorFile = 1

