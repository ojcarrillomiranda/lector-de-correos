import imaplib , email
    
counteraux = 1

    
def inicial():
    print("Llege a LeerCoreosFactura em factura")
    M = imaplib.IMAP4_SSL('172.24.1.98','7993')
    M.login('carga_fe@mct.com.co', 'Mct.2021')
    M.select("Factura")
    ##name_pat = re.compile('name=".*"')
    Msg = "\033[92m Listo para traer mensajes sin leer de la cuenta carga_fe@mct.com.co\033[0m\n"
    print (Msg)
    typ, data = M.search(None, 'UNSEEN')
    for num in data[0].split():
        typ, data = M.fetch(num, '(RFC822)')
        if typ == 'OK':
            raw_email = data[0][1]
            raw_email_string = raw_email.decode('utf-8')
            email_message = email.message_from_string(raw_email_string)
            print(email_message)
inicial()