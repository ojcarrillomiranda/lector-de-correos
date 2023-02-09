#! /usr/bin/env python
# -*- coding: utf-8 -*-
from configparser import ConfigParser
from email.mime.text import MIMEText
import re
from smtplib import SMTP

config = ConfigParser()
config.read('config/config.ini')
host = config.get('enviar_correo', 'SMTP_HOST')
puerto = config.getint('enviar_correo', 'SMTP_PORT')
correo = config.get('enviar_correo', 'SMTP_CORREO')
password = config.get('enviar_correo', 'SMTP_PASSWORD')
email_orlin = config.get('enviar_correo', 'EMAIL_ORLIN')
email_erney = config.get('enviar_correo', 'EMAIL_ERNEY')
email_log = config.get('enviar_correo', 'EMAIL_LOG')
correos = email_orlin + ',' + email_erney + ',' + email_log

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
        mensaje['From'] = correo
        mensaje['To'] = str(email_from)
        mensaje['Subject']="OK " + str(email_message)
        self.enviar_correo(correo,str(email_from),mensaje.as_string())
        carpeta = re.sub("El|cargue|la|del|:","",cargue)
        mensaje = "\033[92mSe envia correo de confirmación de cargue" + str(carpeta.upper()) + " satisfactorio a:\033[0m " + email_from + "\n"
        print(mensaje)

    def correo_error(self, documento, filename, msg, cargue, error, email_from, email_message):
        mensaje = MIMEText(cargue + str(documento) + """ ha generado un ERROR y NO fue realizado.\n\nArchivo: """ + str(filename) + """\n\n""" +msg+ """\n\nCordialmente,\n\n\nEquipo TICs MCT.""")
        mensaje['From']= correo
        mensaje['To']= correos
        mensaje['Subject']= error + str(documento) + " - " + str(email_message)
        self.enviar_correo(correo, [correos],mensaje.as_string())
            # ["facturacion@mct.com.co", "logs.switrans@mct.com.co", "erney.vargas@mct.com.co"],
        carpeta = re.sub("Error|:","",error)
        mensaje = "\033[92mSe envia correo de ERROR de cargue" + carpeta.upper() + " a:\033[0m " + correos
        print(mensaje)
        print("\033[92m._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._\033[0m\n")

    def enviar_correo(self, from_addr, to_addrs, msg):
        self.smtp.sendmail(from_addr, to_addrs, msg)

    def cerrar_servidor(self):
        self.smtp.quit()
        self.smtp.close()
        print("\033[91mConexion a SERVIDOR ENVIO DE CORREOS cerrada con exito\033[0m")