import os
from dotenv import load_dotenv

load_dotenv()

# Configuración de entorno
URL_SIGA = os.getenv('URLSIGA')
URLSECCION = os.getenv('siga_url')
USER_SIGA = os.getenv('UsuarioSIGA')
PASSWORD_SIGA = os.getenv('PASS')
REMITENTE = os.getenv('REMITENTE')
HOST_CORREO = os.getenv('Host')
PUERTO_CORREO = os.getenv('puerto')
PASSWORD_CORREO = os.getenv('PASS_CORREO')
SIGA_USUARIOS= os.getenv('Siga_Usuarios')