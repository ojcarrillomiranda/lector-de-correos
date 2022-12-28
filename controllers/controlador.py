from configparser import ConfigParser
from common.funciones import Funcion
from conexiones.ConexionDB import Conexion
from conexiones.EnviarCorreo import EnviarCorreo
from models.documento_soporte import DocumentoSoporte
from models.factura import Factura
from models.nota_ajuste import NotaAjuste
from models.nota_credito import NotaCredito
from models.nota_debito import NotaDebito


class Controlador(object):
    __instancia = None
    def __new__(cls):
        if Controlador.__instancia is None:
            Controlador.__instancia = object.__new__(cls)
        return Controlador.__instancia

    conexion = Conexion().conexion()
    enviar = EnviarCorreo()

    print("\033[93mEMPEZAMOS EN FACTURA\033[0m")
    Factura().leer_correos_factura(conexion,enviar)

    # print("\033[93mEMPEZAMOS EN NOTA CREDITO\033[0m")
    # NotaCredito().leer_correos_nota_credito(conexion)
    #
    # print("\033[93mEMPEZAMOS EN NOTA DEBITO\033[0m")
    # NotaDebito().leer_correos_nota_debito(conexion)
    #
    # print("\033[93mEMPEZAMOS EN DOCUMENTO SOPORTE\033[0m")
    # DocumentoSoporte().leer_correos_documento_soporte(conexion)
    #
    # print("\033[93mEMPEZAMOS EN NOTA AJUSTE\033[0m")
    # NotaAjuste().leer_correos_nota_ajuste(conexion)
    # obj_factura =Factura(conexion)
    # obj_nota_credito =NotaCredito(conexion)
    # obj_nota_debito =NotaDebito(conexion)
    # obj_documento_soporte =DocumentoSoporte(conexion)
    # obj_nota_ajuste =NotaAjuste(conexion)
    # factura = {
    #     "conexion":conexion,
    #     "enviar": enviar,
    #     "modulo":120,
    #     "imgType":122,
    #     "letra":"E",
    #     "objeto":obj_factura,
    #     "enviar":enviar,
    #     "carpeta":config.get('conf','IMAP_FAC')
    # }
    # nota_credito = {
    #     "conexion":conexion,
    #     "enviar": enviar,
    #     "modulo":129,
    #     "imgType":135,
    #     "letra":"C",
    #     "objeto":obj_nota_credito,
    #     "enviar":enviar,
    #     "carpeta":config.get('conf','IMAP_NC')
    # }
    # nota_debito = {
    #     "conexion":conexion,
    #     "enviar": enviar,
    #     "modulo":256,
    #     "imgType":140,
    #     "letra":"ND",
    #     "objeto":obj_nota_debito,
    #     "enviar":enviar,
    #     "carpeta":config.get('conf','IMAP_ND')
    # }
    # documento_soporte = {
    #     "conexion":conexion,
    #     "enviar": enviar,
    #     "modulo":299,
    #     "imgType":141,
    #     "letra":"D",
    #     "objeto":obj_documento_soporte,
    #     "carpeta":config.get('conf','IMAP_DS')
    # }
    # nota_ajuste = {
    #     "conexion":conexion,
    #     "enviar": enviar,
    #     "modulo":346,
    #     "imgType":285,
    #     "letra":"N",
    #     "objeto":obj_nota_ajuste,
    #     "carpeta":config.get('conf','IMAP_NA')
    # }
    # print("\033[93mEMPEZAMOS EN FACTURA\033[0m")
    # Funcion(factura)
    #
    # print("\033[93mEMPEZAMOS EN NOTA CONTABLE\033[0m")
    # Funcion(nota_credito)
    # print("\033[93mEMPEZAMOS EN NOTA DEBITO\033[0m")
    # Funcion(nota_debito)
    # print("\033[93mEMPEZAMOS EN DOCUMENTO SOPORTE\033[0m")
    # Funcion(documento_soporte)
    # print("\033[93mEMPEZAMOS EN NOTA AJUSTE\033[0m")
    # Funcion(nota_ajuste)
    # NotaCredito().leerCorreosNotaContabilidad(conexion, enviar, 129, 135, "C")
    # NotaDebito().LeerCorreosNotaDebito(conexion, enviar)
    # DocumentoSoporte().leerCorreosDocumentoSoporte(conexion, enviar)
    # NotaAjuste().LeerCorreosNotaAjuste(conexion, enviar)

    conexion.close()
    print("\033[91mConexion DB cerrada con exito\033[0m")
    enviar.cerrar_servidor()
