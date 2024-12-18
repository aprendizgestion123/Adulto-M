import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from Fuji.get_data import GetData
from dotenv import load_dotenv


class Correo:
    def __init__(self): 
        load_dotenv()
        
        """ Data = GetData()
        __data = Data.get_datos_id('4')

        self.SMTP_SERVER = __data['server_smtp']
        self.SMTP_PORT = __data['port_smtp']
        self.SMTP_USERNAME = __data['user_smtp']
        self.SMTP_PASSWORD = __data['pass_smtp'] """

        self.SMTP_SERVER = os.getenv('SMTP_SERVER')
        self.SMTP_PORT = os.getenv('SMTP_PORT')
        self.SMTP_USERNAME = os.getenv('SMTP_USERNAME')
        self.SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')
        # Dirección de correo electrónico del remitente
        self.from_email = self.SMTP_USERNAME
        self.to_email = os.getenv('self.to_email').split(',')

    def crear_mensaje(self, subject, body):
        # Crear el mensaje
        msg = MIMEMultipart()
        msg['From'] = self.from_email
        msg['To'] = ', '.join(self.to_email)
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'html'))
        return msg

    def enviar(self, msg):
        # Iniciar sesión en el servidor SMTP y enviar correo
        try:
            server = smtplib.SMTP(self.SMTP_SERVER, self.SMTP_PORT)
            server.starttls()
            server.login(self.SMTP_USERNAME, self.SMTP_PASSWORD)
            text = msg.as_string()
            server.sendmail(self.from_email, self.to_email, text)
            server.quit()
            print("Correo enviado correctamente")
        except Exception as e:
            print("Error al enviar el correo:", str(e))

