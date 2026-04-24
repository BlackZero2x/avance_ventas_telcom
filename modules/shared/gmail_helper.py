"""
Helper para envío de correos con Gmail API
"""
import logging
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import base64


class GmailHelper:
    def __init__(self, gmail_service):
        self.gmail_service = gmail_service
    
    def send_email_with_attachment(self, recipients, subject, body, attachment_path):
        """
        Envía un correo con archivo adjunto
        
        Args:
            recipients: Lista de emails o string con un email
            subject: Asunto del correo
            body: Cuerpo del correo
            attachment_path: Ruta del archivo a adjuntar
        
        Returns:
            bool: True si se envió correctamente, False si falló
        """
        try:
            logging.info(f"Enviando correo a: {recipients}")
            
            # Convertir a lista si es string
            if isinstance(recipients, str):
                recipients = [recipients]
            
            # Crear mensaje
            message = MIMEMultipart()
            message['to'] = ', '.join(recipients)
            message['subject'] = subject
            
            # Cuerpo del mensaje
            message.attach(MIMEText(body, 'plain'))
            
            # Adjuntar archivo
            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, 'rb') as file:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {os.path.basename(attachment_path)}'
                    )
                    message.attach(part)
            else:
                logging.warning(f"Archivo no encontrado: {attachment_path}")
            
            # Codificar mensaje
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
            
            # Enviar con Gmail API
            send_message = {'raw': raw_message}
            result = self.gmail_service.users().messages().send(
                userId='me',
                body=send_message
            ).execute()
            
            logging.info(f"Correo enviado exitosamente. ID: {result['id']}")
            return True
            
        except Exception as e:
            logging.error(f"Error enviando correo: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False
    
    def send_simple_email(self, recipients, subject, body):
        """
        Envía un correo simple sin adjuntos
        
        Args:
            recipients: Lista de emails o string con un email
            subject: Asunto del correo
            body: Cuerpo del correo
        
        Returns:
            bool: True si se envió correctamente, False si falló
        """
        try:
            logging.info(f"Enviando correo simple a: {recipients}")
            
            if isinstance(recipients, str):
                recipients = [recipients]
            
            message = MIMEMultipart()
            message['to'] = ', '.join(recipients)
            message['subject'] = subject
            message.attach(MIMEText(body, 'plain'))
            
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
            send_message = {'raw': raw_message}
            
            result = self.gmail_service.users().messages().send(
                userId='me',
                body=send_message
            ).execute()
            
            logging.info(f"Correo enviado exitosamente. ID: {result['id']}")
            return True
            
        except Exception as e:
            logging.error(f"Error enviando correo: {e}")
            return False