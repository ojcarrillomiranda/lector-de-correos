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
    Factura().leer_correos_factura(conexion, enviar)

    print("\033[93mEMPEZAMOS EN NOTA CONTABLE\033[0m")
    NotaCredito().leer_correos_nota_credito(conexion, enviar)

    print("\033[93mEMPEZAMOS EN NOTA DEBITO\033[0m")
    NotaDebito().leer_correos_nota_debito(conexion, enviar)

    print("\033[93mEMPEZAMOS EN DOCUMENTO SOPORTE\033[0m")
    DocumentoSoporte().leer_correos_documento_soporte(conexion, enviar)

    print("\033[93mEMPEZAMOS EN NOTA AJUSTE\033[0m")
    NotaAjuste().leer_correos_nota_ajuste(conexion, enviar)

    conexion.close()
    print("\033[91mConexion DB cerrada con exito\033[0m")
    enviar.cerrar_servidor()
