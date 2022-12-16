from datetime import datetime as dater
import datetime
import imaplib
import email
from email.policy import default as default_policy
from bs4 import BeautifulSoup as b
import yaml
import pandas as pd
from emailsender import senderEmail
from whatsapi import whatsappSender

#List that save the mail direction to not repeat the message
list_sent = []
#List that save the whatsapp number to not repeat the message
lista_enviados_whatsapp = []

#Open the file where the mails alredy sended save, and save into a list to acces inmediately
with open ('emailUserSent.txt', 'rt') as myfile:
    for myline in myfile:
        esteSi = myline.replace(' ', '').replace('\n', '')
        list_sent.append(esteSi)


#Open the file where the whastsapp numbers alredy sended save, and save into a list to acces inmediately
with open ('whatsappEnviados.txt', 'rt') as my_whatsapp_file:
    for myline in my_whatsapp_file:
        whatsappEnviados = myline.replace(' ', '').replace('\n', '')
        lista_enviados_whatsapp.append(whatsappEnviados)
        
#Open the fail to appen new data to both files 
opener = open(r'emailUserSent.txt', 'a+')
openerWhatsapp = open(r'whatsappEnviados.txt', 'a+')

#Function that save the mails direction of the leads on the file
def saver(theElement):
    nuevo = theElement.replace(' ', '')
    opener.write(nuevo)
    opener.write('\n')

#Function that save the whatsapp numbers of the leads on the file
def Whatsapp_saver(theElement):
    nuevo_numero = theElement.replace(' ', '')
    openerWhatsapp.write(nuevo_numero)
    openerWhatsapp.write('\n')







##Variables to make the dataFrame
# This is where all clients are saved
all_table = {}
# This is where the clients that took the insurance are saved
all_table_new_Client = {}


# This is the function that recolect all the mails 
def mailColecter():
    # Open the file where the password and the user are saved
    with open("passwordMine.yml") as f:
        content = f.read()
    
    my_credentials = yaml.load(content, Loader=yaml.FullLoader)

    user, password = my_credentials["user"], my_credentials["password"]

    imap_url = "imap.gmail.com"


    my_mail = imaplib.IMAP4_SSL(imap_url)

    my_mail.login(user, password)

    my_mail.select('Inbox')


    date = (datetime.date.today() - datetime.timedelta(days=2)).strftime("%d-%b-%Y")

    typ, data = my_mail.search(None, '(ALL)', f'(SENTSINCE {date})')

    mail_id_list = data[0].split()


    mesages = []

    counter = 0
    counter_client = 0

    for num in mail_id_list:
        typ, data = my_mail.fetch(num, '(RFC822)')
        mesages.append(data)


    for msg in mesages[::-1]:
    #for msg in mesages[-10:]:
        for response_part in msg:
            if type(response_part) is tuple:
                my_mesage = email.message_from_bytes((response_part[1]), policy=default_policy)

                if my_mesage['subject'] == 'Poliza nueva creada':
                    for part in my_mesage.walk():
                        if part.get_content_type() == 'text/html':
                            body = part.get_body(preferencelist=('plain', 'html'))

                            maybe_decoded_payload = body.get_payload(decode=True)
                            if (maybe_decoded_payload is not None):
                                html = bytes.decode(maybe_decoded_payload, encoding="utf-8")
                                soup = b(html, features="html.parser")
                                tds = soup.find_all('span')
                                csv_data = []

                                for td in tds:
                                    inner_text = td.text
                                    strings = inner_text.split("\n")

                                    csv_data.extend([string for string in strings if string])

                                (",".join(csv_data))         
                                my_list = csv_data[2].split(':')
                                my_list.remove('Este correo es para notificarte que se acaba de registrar un nuevo intento de cotización.Nombre')
                                
                                nombreCliente = my_list[0].replace('Apellido', '').replace(' ', '')
                                apellidoCliente= my_list[1].replace('N. identificación', '').replace(' ', '')
                                correo = my_list[3].replace('Teléfono', '').replace(' ', '')
                                celularCliente= my_list[4].replace('Fecha nacimiento', '')
                                phoneClient = ('57' + str(celularCliente)).replace(' ', '')

                                cedula = my_list[2].replace('Correo', '').replace(' ', '')
                                fecha_nacimiento = my_list[5].replace('Genero', '').replace(' ', '')
                                genero = my_list[6].replace('Ciudad', '').replace(' ', '')
                                ciudad = my_list[7].replace('Placa', '').replace(' ', '')
                                placaCliente = my_list[8].replace('Modelo', '').replace(' ', '')
                                modelo =  my_list[9].replace('Valor facecolda', '').replace(' ', '')
                                valor_fasecolda = my_list[10].replace('https', '').replace(' ', '')
                                mensaje_cotizar = f"Perfecto {nombreCliente}!, te confirmo tus datos, cotización para el vehículo con placas: {placaCliente}, modelo: {modelo}, circulando en {ciudad}. Nombre del titular: {nombreCliente}, cédula del titular: {cedula}, fecha de nacimiento del titular: {fecha_nacimiento} "

                                my_last_dictionary = {
                                    'Nombre': nombreCliente,
                                    'Apellido': apellidoCliente,
                                    'Correo': correo, 
                                    'Telefono': int(phoneClient), 
                                    'Cedula': cedula,
                                    'Fecha_naciemnto': fecha_nacimiento,
                                    'Genero': genero,
                                    'Ciudad': ciudad,
                                    'Placa': placaCliente, 
                                    'Valor_fasecolda': valor_fasecolda,
                                    'Modelo': modelo,
                                    'DateCreated': dater.now(),
                                    'mensaje': mensaje_cotizar

                                }
                                counter += 1

                                if correo not in list_sent:
                                    senderEmail(nombreCliente, correo)
                                    #Here the mail is sent 
                                    saver(correo)
                                    list_sent.append(correo)
                                    print('Email enviado con exito a ' + correo)

                                    
                                else: 
                                    print('Correo repetido')
                                    pass



                                if phoneClient not in lista_enviados_whatsapp:
                                    Whatsapp_saver(phoneClient)
                                    lista_enviados_whatsapp.append(phoneClient)
                                    #Here the whatsapp message is sent to the lead
                                    whatsappSender(nombreCliente, phoneClient)
                                    print('numero guardado')

                                else:
                                    print('Numero repetido')
                            

                                if len(str(my_last_dictionary['Telefono'])) != 12:
                                    print("Numero no valido")
                                    pass
                                
                                else:    
                                    all_table[counter] = my_last_dictionary


                elif my_mesage['subject'] == 'Nueva cotización':
                    for part in my_mesage.walk():
                        if part.get_content_type() == 'text/html':
                            body = part.get_body(preferencelist=('plain', 'html'))

                            maybe_decoded_payload = body.get_payload(decode=True)
                            if (maybe_decoded_payload is not None):
                                html = bytes.decode(maybe_decoded_payload, encoding="utf-8")
                                soup = b(html, features="html.parser")
                                tds = soup.find_all('span')
                                csv_data_clientes_potenciales = []

                                for td in tds:
                                    inner_text = td.text
                                    strings = inner_text.split("\n")

                                    csv_data_clientes_potenciales.extend([string for string in strings if string])

                                (",".join(csv_data_clientes_potenciales))

                                my_list_client = csv_data_clientes_potenciales
                         

    
                                nombre = my_list_client[4].replace('Nombre:', '')
                                celular= my_list_client[6].replace('Celular:', '').replace('\r:', '').replace(' ', '')
                                correo = my_list_client[8].replace('Correo:', '').replace(' ', '')
                                placa= my_list_client[10].replace('Placa:', '')
                                valor= my_list_client[12].replace('Valor:', '')
                                aseguradora= my_list_client[18].replace('Aseguradora:', '')
                                phone = '57' + str(celular)


                            my_last_dictionary_client = {
                                'Nombre':  nombre,
                                'Celular': int(phone),
                                'Correo': correo, 
                                'Placa': placa,
                                'Valor': valor,
                                'Aseguradora': aseguradora,
                                'DateCreated': dater.now()

                            }

                            counter_client += 1

                            if len(str(my_last_dictionary_client['Celular'])) != 12:
                                print("Numero no valido")
                                pass
                            
                            else:
                                all_table_new_Client[counter_client] = my_last_dictionary_client



def allTheClients():
    df_all_clients = pd.read_excel('clients.xlsx')
    new_dats = pd.DataFrame(all_table)
    new_dats = new_dats.transpose()
    new_dataframe = pd.concat([df_all_clients, new_dats])
    new_dataframe = new_dataframe.drop_duplicates(subset='Telefono')
    new_dataframe = new_dataframe.drop_duplicates(subset='Correo')
    new_dataframe = new_dataframe[['Nombre', 'Apellido', 'Correo', 'Telefono', 'Cedula', 'Fecha_naciemnto', 'Genero', 'Ciudad',  'Placa', 'Valor_fasecolda', 'Modelo', 'mensaje', 'DateCreated']]
    new_dataframe = new_dataframe.sort_values(by=['DateCreated'], ascending=False)
    new_dataframe.to_excel('clients.xlsx',  index= False)
    print('All clients succesfully added ')
    print('________________________________________________________________________________')

def potentialClients():
    df_all_potentials =  pd.read_excel('potential.xlsx')
    new_potentials = pd.DataFrame(all_table_new_Client)
    new_potentials = new_potentials.transpose()
    new_potential_dataframe = pd.concat([df_all_potentials, new_potentials])
    new_potential_dataframe =new_potential_dataframe.drop_duplicates(subset='Celular')
    new_potential_dataframe =new_potential_dataframe.drop_duplicates(subset='Celular')
    new_potential_dataframe= new_potential_dataframe.sort_values(by=['DateCreated'], ascending=False)
    new_potential_dataframe.to_excel('potential.xlsx', index= False)
    print("Potential clients succesfully added")


def main():
    mailColecter()
    allTheClients()
    potentialClients()
    opener.close()
    openerWhatsapp.close()

    hoy = dater.now().strftime("%d/%m/%Y %H:%M:%S")
    print("Last time added" + " " +  hoy)

if __name__ == "__main__":
    main()
    

