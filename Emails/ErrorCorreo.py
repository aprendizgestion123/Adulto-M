from Emails.correo import Correo


class ErrorCorreo(Correo):
    def __init__(self):
        super().__init__()

    def enviar_error(self,mensaje):
        # Contenido del correo (puede incluir HTML)
        body = f"""
        <html>
        <body>
            <p>Cordial saludo,<br>
            <br>
            Ocurrió un error en el proceso de Ingreso vendedores Externos.<br>
            <br>
            Detalles del error: <br><br><b>{mensaje}</b><br>
            <br>
            Por favor no responder ni enviar correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com.
            </p>
        </body>
        </html>
        """
        subject = 'Error en el proceso de Ingreso vendedores Externos'
        msg = super().crear_mensaje(subject, body)
        super().enviar(msg)  # Llamar al método de la clase padre