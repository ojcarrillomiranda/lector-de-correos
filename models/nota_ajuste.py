#! /usr/bin/env python
# -*- coding: utf-8 -*-
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
carpeta = config.get('doc_soporte','IMAP_NA')
cargue_string = config.get('correo_confirmacion','CARGUE_NOTA_AJUSTE')
error_string = config.get('correo_error','ERROR_NOTA_AJUSTE')

class NotaAjuste:

    def leer_correos_nota_ajuste(self, con, env):
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

                valida_asunto = Funcion(cursor, conexion, enviar).validar_asunto_correo("N", asunto)
                empresa_codigo = valida_asunto["empresa_codigo"]
                documento = valida_asunto["cod_doc_asunto"]
                empresa = valida_asunto["empresa"]
                if valida_asunto is False:
                    mensaje = "\033[91mEl Asunto no es correcto o esta mal formado (no tiene el numero de documento soporte), se envía correo y se detiene el proceso\033[0m"
                    print(mensaje)
                    break
                datos_nota_ajuste = self.get_datos_nota_ajuste(cursor, empresa_codigo, documento, empresa)
                nota_ajuste = datos_nota_ajuste["cod_nota_ajuste"]
                ruta = datos_nota_ajuste["digitalizado_ruta"]
                if datos_nota_ajuste is False:
                    mensaje = "\033[91No se lograron obtener datos de la " + cambio.upper() + "\033[0m"
                    print(mensaje)
                    break
                insert_digi = Funcion(cursor, conexion, enviar).insertar_digitalizado(346, 285, nota_ajuste, ruta, filename, cargue_string, error_string)
                if insert_digi:
                    archivo_zip = insert_digi
                    self.actualizar_nota_ajuste(cursor, conexion, archivo_zip, nota_ajuste)
                    Funcion(cursor, conexion, enviar).eventos(nota_ajuste, 346)
            mensaje = "\033[91mNo existen mensajes sin leer en la cuenta " + correo + " en la carpeta " + cambio.upper() + "\033[0m"
            print (mensaje)
            mensaje = "\033[93m*********************************************** Proceso Finalizado " + cambio.upper() + " ***********************************************\033[0m\n\n"
            print (mensaje)

    def get_datos_nota_ajuste(self, cursor, empresa_codigo, nota_ajuste, empresa):
        global detach_dir, detach_dir1, cod_nota_ajuste, cencos_nota_ajuste, fecha_nota_ajuste, ano, mes, dia
        nota_ajuste_sql = "SELECT na.notaju_codigo, na.notaju_fechacreacion, na. docsop_codigo, na.empresa_codigo, na.cencos_codigo FROM tb_notaajuste na WHERE na.empresa_codigo = " + str(empresa_codigo) + " AND na.notaju_cencoscodigo = "+ str(nota_ajuste)
        mensaje = "\033[94mConsulta de la Nota Ajuste de la Empresa " + str(empresa_codigo) + " DigCCResolucion " + str(nota_ajuste) + " ejecutada con éxito\033[0m"
        try:
            print("\033[94m######################### SQL get_datos() #################################\033[0m")
            print ("\033[94m" + nota_ajuste_sql + "\033[0m")
            cursor.execute(nota_ajuste_sql)
            print (mensaje)
            outfile = '.'
            ContAux = 0
            for row in cursor.fetchall():
                r = reg(cursor, row)
                cod_nota_ajuste = r.notaju_codigo
                cencos_nota_ajuste = r.cencos_codigo
                fecha_nota_ajuste = str(r.notaju_fechacreacion)
                ano = fecha_nota_ajuste.split("-")[0]
                mes = fecha_nota_ajuste.split("-")[1]
                dia = fecha_nota_ajuste.split("-")[2]
                digitalizado_ruta = '/usr/local/apache/htdocs/switrans/images/documentosoporte/' + str(empresa) + '/' + str(
                    ano) + '/' + str(cencos_nota_ajuste) + '/' + str(mes) + '/'
                ContAux = ContAux + 1
                print("\033[92m######################### RESULTADO SQL get_datos() #################################\033[0m\n" + "\033[92mnotaju_codigo = " + str(
                    cod_nota_ajuste) + "\n" + "cencos_codigo = " + str(cencos_nota_ajuste) + "\n" + "notaju_fechacreacion = " + fecha_nota_ajuste + "\033[0m\n")
            if ContAux > 0:
                datos_nota_ajuste = {
                    "cod_nota_ajuste": cod_nota_ajuste,
                    "digitalizado_ruta": digitalizado_ruta
                }
                return datos_nota_ajuste
            else:
                return False
        except psycopg2.Error as e:
            print(e.pgerror)
        return False

    def actualizar_nota_ajuste(self, cursor, conexion, filename, nota_ajuste):
        updateNotaAjuste = (
                "UPDATE tb_notaajuste SET archivo_notaajuste_electronico = '" + filename + "' WHERE notaju_codigo = " + str(
            nota_ajuste))
        try:
            cursor.execute(updateNotaAjuste)
            conexion.commit()
            mensaje = "\033[93mNota Ajuste Actualizada a base de datos tb_notaajuste, campo: archivo_notaajuste_electronico \n._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._\033[0m"
            print(mensaje)
        except psycopg2.Error as e:
            mensaje = e.pgerror
            print(mensaje)
            print("ERROR FATAL EN LA CONSULTA " + str(updateNotaAjuste))
            errorFile = 1