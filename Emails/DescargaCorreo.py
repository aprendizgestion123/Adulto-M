import os
from datetime import datetime
from O365 import Account, FileSystemTokenBackend
from os.path import exists
from dotenv import load_dotenv
import re

class DescargaCorreo:
    def __init__(self):
        
        load_dotenv()

        self.path_ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        """  self.db_siga_excel = os.path.join(self.path_, 'Insumo', 'plantillaParaActivacion.xlsx')
        self.path_procesados = os.path.join(self.path_, 'Historicos', 'PlantillasActivacionProcesadas')
        self.path_procesados_mes = os.path.join(self.path_procesados, f'{datetime.now().strftime("%Y-%B")}')
        self.download_path = os.path.join(self.path_, 'Insumo')
        self.file_name_template = "plantillaParaActivacion.xlsx" """
        # Cargar variables de entorno
    
        """  Data = GetData()
        __data = Data.get_datos_id('4')
        self.CLIENT_ID = __data['client_id']
        self.CLIENT_SECRET = __data['secret_id']
        self.TENANT_ID = __data['tenant_id'] """
        self.CLIENT_ID = os.getenv('CLIENT_ID')
        self.CLIENT_SECRET = os.getenv('CLIENT_SECRET')
        self.TENANT_ID = os.getenv('TENANT_ID')

    
    def obtener_fechas_desde_correo(self):
        try:
            # Configuración y autenticación
            credentials = (self.CLIENT_ID, self.CLIENT_SECRET)
            token_path = os.path.join(self.path_, 'o365_token.txt')
            token_backend = FileSystemTokenBackend(token_path=token_path)
            account = Account(credentials, tenant_id=self.TENANT_ID, token_backend=token_backend)

            if not account.is_authenticated:
                if account.authenticate(scopes=['basic', 'message_all']):
                    print('Autenticado correctamente y token almacenado.')
                else:
                    print('Error de autenticación')
                    return False, "Error de autenticación",None

            # Obtener correos
            messages = list(account.mailbox().get_folder(folder_name='Adulto Mayor').get_messages(
                query="isRead eq false and subject eq 'Fechas Adulto Mayor'", download_attachments=True))
            
            if not messages:
                print("En la bandeja de entrada NO hay correos nuevos")
                return False, "No hay correos nuevos",None

            # Procesar el último correo
            last_message = messages[-1]
            Cuerpo = last_message.get_body_text()
            
            # Extraer fechas con expresión regular
            fechas_ = r'\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{2}-\d{2}|\d{1,2}\sde\s(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\sde\s\d{4})\b'
            fechas = re.findall(fechas_, Cuerpo)

            if fechas:
                # Detectar y ordenar fechas
                def parse_fecha(fecha):
                    formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]
                    for fmt in formatos:
                        try:
                            return datetime.strptime(fecha,fmt)
                        except ValueError:
                            continue
                    raise ValueError(f"Formato no reconocido para la fecha: {fecha}")

                # Ordenar las fechas extraídas
                fechas = sorted(fechas, key=parse_fecha)
                fecha_inicio = fechas[0]
                fecha_fin = fechas[1] if len(fechas) > 1 else None
                fecha_inicio = datetime.strptime(fecha_inicio, '%d/%m/%Y')
                fecha_fin = datetime.strptime(fecha_fin, '%d/%m/%Y')
                fecha_inicio= fecha_inicio.strftime('%d-%m-%Y')
                fecha_fin= fecha_fin.strftime('%d-%m-%Y')
                print(fecha_inicio)
                # Marcar el correo como leído
                last_message.mark_as_read()

                # Retornar las fechas
                return True,fecha_inicio, fecha_fin
            else:
                print("No se encontraron fechas en el cuerpo del mensaje")
                return False, "No se encontraron fechas",None

        except Exception as e:
            print(f"Error: {e}")
            return False, f"Error inesperado: {e}",None
