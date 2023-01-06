#! /usr/bin/env python
# -*- coding: utf-8 -*-
from configparser import ConfigParser
from email.mime.text import MIMEText
from fileinput import close
import re
from smtplib import SMTP

config = ConfigParser()
config.read('config/config.ini')
host = config.get('enviar_correo', 'SMTP_HOST')
puerto = config.getint('enviar_correo', 'SMTP_PORT')
correo = config.get('enviar_correo', 'SMTP_CORREO')
password = config.get('enviar_correo', 'SMTP_PASSWORD')

class EnviarCorreo(object):
    __instancia = None
    def __new__(cls):
        if EnviarCorreo.__instancia is None:
            EnviarCorreo.__instancia = object.__new__(cls)
        return EnviarCorreo.__instancia
    def __init__(self):
        self.smtp = SMTP()
        self.smtp.connect(host, puerto)
        self.smtp.starttls()
        self.smtp.login(correo, password)

    def correo_confirmacion(self, cod_documento, cargue, email_from, email_message):
        mensaje = MIMEText(cargue + str(cod_documento) + """  fue realizado correctamente. Puede ingresar a Switrans y verificar.\n\nCordialmente,\n\n\nEquipo TICs MCT.""")
        mensaje['From']="notificaciones@mct.com.co"
        mensaje['To']= str(email_from)
        mensaje['Subject']="OK " + str(email_message)
        self.enviar_correo("notificaciones@mct.com.co",str(email_from),mensaje.as_string())
        carpeta = re.sub("El|cargue|de|la|del|:","",cargue)
        mensaje = "\033[92mSe envia correo de confirmación de cargue" + str(carpeta.upper()) + " satisfactorio a:\033[0m " + email_from + "\n"
        print(mensaje)

    def correo_error(self, factura, filename, msg, cargue, error, email_from, email_message):
        c = "orlin.carrillo@mct.com.co"
        mensaje = MIMEText(cargue + str(factura) + """ ha generado un ERROR y NO fue realizado.\n\nArchivo: """ + str(filename) + """\n\n""" +msg+ """\n\nCordialmente,\n\n\nEquipo TICs MCT.""")
        mensaje['From']="notificaciones@mct.com.co"
        mensaje['To']= str(email_from) + ",logs.switrans@mct.com.co"
        mensaje['Subject']=error + str(factura) + " - " + str(email_message)
        self.enviar_correo("notificaciones@mct.com.co", c,mensaje.as_string())
            # ["facturacion@mct.com.co", "logs.switrans@mct.com.co", "erney.vargas@mct.com.co"],
        carpeta = re.sub("Error|:","",error)
        mensaje = "\033[92mSe envia correo de ERROR de cargue" + carpeta.upper() + " a:\033[0m " + c
        print(mensaje)
        print("\033[92m._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._\033[0m\n")

    def enviar_correo(self, from_addr, to_addrs, msg):
        self.smtp.sendmail(from_addr, to_addrs, msg)

    def cerrar_servidor(self):
        self.smtp.quit()
        self.smtp.close()
        print("\033[91mConexion a SERVIDOR ENVIO DE CORREOS cerrada con exito\033[0m")