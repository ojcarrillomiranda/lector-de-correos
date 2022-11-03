#! /usr/bin/env python
# -*- coding: utf-8 -*-
from configparser import ConfigParser
from email.mime.text import MIMEText
from fileinput import close
import re
# import smtplib 
from smtplib import SMTP 

config = ConfigParser()
config.read('config/config.ini')
class EnviarCorreo(object):
    __instancia = None
    def __new__(cls):
        if EnviarCorreo.__instancia is None:
            EnviarCorreo.__instancia = object.__new__(cls)
        return EnviarCorreo.__instancia
    def __init__(self):
        self.smtp = SMTP()
        self.smtp.connect(config.get('conf', 'SMTP_HOST'), config.getint('conf', 'SMTP_PORT'))
        self.smtp.starttls()
        self.smtp.login(config.get('conf', 'SMTP_CORREO'), config.get('conf', 'SMTP_PASSWORD'))

    def correoConfirmacion(self, cod_documento, modulo, emailFrom, email_message):
        modulo_string = ""
        cargue = ""
        if modulo == 120:
            modulo_string = "la factura:"
            cargue = "El cargue de la factura"
        elif modulo == 129:
            modulo_string = "la nota credito:"
            cargue = "El cargue de la nota credito"
        elif modulo == 256:
            modulo_string = "la nota debito:"
            cargue = "El cargue de la nota debito"
        elif modulo == 299:
            modulo_string = "el documento soporte:"
            cargue = "El cargue del documento soporte"
        elif modulo == 346:
            modulo_string = "la nota ajuste:"
            cargue = "El cargue de la nota ajuste"
        mensaje = MIMEText(cargue + """ para """ + str(modulo_string) + str(cod_documento) + """  fue realizado correctamente. Puede ingresar a Switrans y verificar.\n\nCordialmente,\n\n\nEquipo TICs MCT.""")
        mensaje['From']="notificaciones@mct.com.co"
        mensaje['To']= str(emailFrom)
        mensaje['Subject']="OK " + str(email_message)
        self.enviar_correo("notificaciones@mct.com.co",str(emailFrom),mensaje.as_string())
        carpeta = re.sub("la|el|:","",modulo_string)
        Msg = "\033[92menviando correo de confirmación de cargue" + carpeta.upper() + " satisfactorio a:\033[0m " + emailFrom + "\n"
        print (Msg)

    def correoError(self, cod_documento, filename, msg, modulo, emailFrom, email_message):
        c = "orlin.carrillo@mct.com.co"
        modulo_string = ""
        error = ""
        if modulo == 120:
            modulo_string = "El cargue de la factura:"
            error = "Error Factura: "
        elif modulo == 129:
            modulo_string = "El cargue de la nota credito:"
            error = "Error Nota credito: "
        elif modulo == 256:
            modulo_string = "El cargue de la nota debito:"
            error = "Error Nota debito: "
        elif modulo == 299:
            modulo_string = "El cargue del documento soporte:"
            error = "Error Documento soporte: "
        elif modulo == 346:
            modulo_string = "El cargue de la nota ajuste:"
            error = "Error Nota ajuste: "
        mensaje = MIMEText(modulo_string + str(cod_documento) + """ ha generado un ERROR y NO fue realizado.\n\nArchivo: """ + str(filename) + """\n\n""" +msg+ """\n\nCordialmente,\n\n\nEquipo TICs MCT.""")
        mensaje['From']="notificaciones@mct.com.co"
        mensaje['To']= str(emailFrom) + ",logs.switrans@mct.com.co"
        mensaje['Subject']=error + str(cod_documento) + " - " + str(email_message)
        self.enviar_correo("notificaciones@mct.com.co", c,mensaje.as_string())
            # ["facturacion@mct.com.co", "logs.switrans@mct.com.co", "erney.vargas@mct.com.co"],
        carpeta = re.sub("Error|:","",error)
        Msg = "\033[92menviando correo de ERROR de cargue" + carpeta.upper() + "a:\033[0m " + c + "\n"
        print (Msg)

    def enviar_correo(self, from_addr, to_addrs, msg):
        self.smtp.sendmail(from_addr, to_addrs, msg)

    def cerrar_servidor(self):
        self.smtp.quit()
        self.smtp.close()
        print("\033[91mConexion a SERVIDOR ENVIO DE CORREOS cerrada con exito\033[0m")