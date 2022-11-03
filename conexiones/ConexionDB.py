#! /usr/bin/env python
# -*- coding: utf-8 -*-
import datetime
import psycopg2.extras
import psycopg2.extensions
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)
from configparser import ConfigParser

datenow = datetime.datetime.now()
hournow = str(datenow.hour) + ":" + str(datenow.minute) + ":" + str(datenow.second)

class Conexion(object):
    __instancia = None
    def __new__(cls):
        if Conexion.__instancia is None:
            Conexion.__instancia = object.__new__(cls)
        return Conexion.__instancia
    def conexion(self):
        config = ConfigParser()
        config.read('config/config.ini')
        try:
            conexion = psycopg2.connect(
                    host = config.get('conf', 'HOST'),
                    dbname = config.get('conf', 'DB_NAME'),
                    user = config.get('conf', 'USER'),
                    password = config.get('conf', 'PASSWORD')
                )
            print("\033[092m#######################################\033[0m\n\033[092m# conexion a DB establecida con exito #" + "\n      " + str(datenow) + " \033[0m\n\033[092m#######################################\033[0m\n")
            return conexion
        except psycopg2.Error as e:
            print ('Error en la conexion de la db!')
            print (e.pgerror)
            print (e.diag.message_detail)
