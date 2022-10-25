from Modelos.Factura import Factura
class Controlador():

    def LlamadoClasesLeidaCorreos(self):
        print("EL ALGOTIRMO LLEGO ACA LlamadoClasesLeidaCorreos")
        Factura().LeerCorreosFactura()
        
        return '\n'
