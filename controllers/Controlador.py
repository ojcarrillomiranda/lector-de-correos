from conexiones.ConexionDB import Conexion
from conexiones.EnviarCorreo import EnviarCorreo
from models.DocumentoSoporte import DocumentoSoporte
from models.Factura import Factura
from models.NotaAjuste import NotaAjuste
from models.NotaContabilidad import NotaContabilidad
from models.NotaDebito import NotaDebito
class Controlador(object):
    __instancia = None
    def __new__(cls):
        if Controlador.__instancia is None:
            Controlador.__instancia = object.__new__(cls)
        return Controlador.__instancia
    def LlamadoClasesLeidaCorreos(self):
        conexion = Conexion().conexion()
        enviar = EnviarCorreo()
        
        Factura().LeerCorreosFactura(conexion, enviar)
        NotaContabilidad().leerCorreosNotaContabilidad(conexion, enviar)
        NotaDebito().LeerCorreosNotaDebito(conexion, enviar)
        DocumentoSoporte().leerCorreosDocumentoSoporte(conexion, enviar)
        NotaAjuste().LeerCorreosNotaAjuste(conexion, enviar)
        
        conexion.close()
        print("\033[91mConexion DB cerrada con exito\033[0m") 
        enviar.cerrar_servidor()
