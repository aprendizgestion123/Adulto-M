import os
from email.mime.base import MIMEBase
from email import encoders
from Emails.correo import Correo

class EnvioCorreo(Correo):
    def __init__(self):
        super().__init__()
        self.path_ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.file_path = os.path.join(self.path_, 'Insumo', 'ADULTO MAYOR ESTANDAR.xlsx')
    
    def adjuntar_archivo(self, msg):
        # Adjuntar el archivo al mensaje
        try:
            file_name = os.path.basename(self.file_path)
            with open(self.file_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {file_name}",
                )
                msg.attach(part)
        except Exception as e:
            print(f"Error al adjuntar el archivo: {e}")
    
    def enviar_correo(self, mensaje):
        # Crear el cuerpo del correo
        body = f"""
        <html>
        <body>
            <p>Cordial saludo,<br>
            <br>
            Se envía el reporte de adulto mayor <br><b>
            {mensaje}</b><br>
            <br>
            Por favor no responder ni enviar correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com.
            </p>
        </body>
        </html>
        """
        subject = 'Adulto Mayor'
        msg = super().crear_mensaje(subject, body)
        self.adjuntar_archivo(msg)  # Adjuntar el archivo al mensaje
        super().enviar(msg)  # Llamar al método de la clase padre para enviar el correo