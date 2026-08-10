"""
Helper para envío de correos con Gmail API
"""
import logging
import os
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import base64


def _enviar_con_reintento(send_fn, intentos=3, delay=15):
    """Ejecuta send_fn() con hasta `intentos` reintentos ante errores de red."""
    for i in range(intentos):
        try:
            return send_fn()
        except Exception as e:
            es_red = any(k in str(e) for k in ("10053", "10054", "ConnectionAborted",
                                                "ConnectionReset", "RemoteDisconnected",
                                                "BrokenPipe", "Connection reset"))
            if es_red and i < intentos - 1:
                logging.warning(f"Error de red al enviar correo (intento {i+1}/{intentos}): {e}. Reintentando en {delay}s...")
                time.sleep(delay)
            else:
                raise


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
            send_message = {'raw': raw_message}
            svc = self.gmail_service
            result = _enviar_con_reintento(
                lambda: svc.users().messages().send(userId='me', body=send_message).execute()
            )
            logging.info(f"Correo enviado exitosamente. ID: {result['id']}")
            return True
            
        except Exception as e:
            logging.error(f"Error enviando correo: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False
    
    def send_email_with_image_body(self, recipients, subject, html_body, image_path, attachment_path=None):
        """
        Envía un correo HTML con una imagen incrustada en el cuerpo (Content-ID)
        y opcionalmente un archivo adjunto adicional.

        Args:
            recipients: lista de emails o string
            subject: asunto
            html_body: HTML con <img src="cid:tds_image"> donde quieras la imagen
            image_path: ruta local de la imagen a incrustar
            attachment_path: ruta opcional de adjunto adicional

        Returns:
            bool
        """
        try:
            if isinstance(recipients, str):
                recipients = [recipients]

            outer = MIMEMultipart("related")
            outer["to"] = ", ".join(recipients)
            outer["subject"] = subject

            alt = MIMEMultipart("alternative")
            outer.attach(alt)
            alt.attach(MIMEText(html_body, "html", "utf-8"))

            if image_path and os.path.exists(image_path):
                with open(image_path, "rb") as f:
                    img = MIMEImage(f.read())
                img.add_header("Content-ID", "<tds_image>")
                img.add_header("Content-Disposition", "inline", filename=os.path.basename(image_path))
                outer.attach(img)
            else:
                logging.warning(f"Imagen no encontrada: {image_path}")

            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(attachment_path)}",
                )
                outer.attach(part)

            raw = base64.urlsafe_b64encode(outer.as_bytes()).decode()
            svc = self.gmail_service
            result = _enviar_con_reintento(
                lambda: svc.users().messages().send(userId="me", body={"raw": raw}).execute()
            )
            logging.info(f"Correo con imagen enviado. ID: {result['id']}")
            return True

        except Exception as e:
            logging.error(f"Error enviando correo con imagen: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def send_email_with_two_images(self, recipients, subject, html_body,
                                    image_path_1=None, image_path_2=None,
                                    attachment_path=None):
        """
        Correo HTML con hasta dos imágenes incrustadas (cid:tds_image_1 y cid:tds_image_2)
        y un adjunto opcional.
        """
        try:
            if isinstance(recipients, str):
                recipients = [recipients]

            outer = MIMEMultipart("related")
            outer["to"] = ", ".join(recipients)
            outer["subject"] = subject

            alt = MIMEMultipart("alternative")
            outer.attach(alt)
            alt.attach(MIMEText(html_body, "html", "utf-8"))

            for img_path, cid in [(image_path_1, "tds_image_1"), (image_path_2, "tds_image_2")]:
                if img_path and os.path.exists(img_path):
                    with open(img_path, "rb") as f:
                        img = MIMEImage(f.read())
                    img.add_header("Content-ID", f"<{cid}>")
                    img.add_header("Content-Disposition", "inline", filename=os.path.basename(img_path))
                    outer.attach(img)
                elif img_path:
                    logging.warning(f"Imagen no encontrada: {img_path}")

            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(attachment_path)}",
                )
                outer.attach(part)

            raw = base64.urlsafe_b64encode(outer.as_bytes()).decode()
            svc = self.gmail_service
            result = _enviar_con_reintento(
                lambda: svc.users().messages().send(userId="me", body={"raw": raw}).execute()
            )
            logging.info(f"Correo con dos imagenes enviado. ID: {result['id']}")
            return True

        except Exception as e:
            logging.error(f"Error enviando correo con imagenes: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def send_email_html(self, recipients, subject, html_body):
        """
        Envía un correo con cuerpo HTML, sin adjuntos.

        Args:
            recipients: lista de emails o string
            subject: asunto
            html_body: contenido HTML del cuerpo

        Returns:
            bool
        """
        try:
            if isinstance(recipients, str):
                recipients = [recipients]

            message = MIMEMultipart("alternative")
            message["to"] = ", ".join(recipients)
            message["subject"] = subject
            message.attach(MIMEText(html_body, "html", "utf-8"))

            raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
            svc = self.gmail_service
            result = _enviar_con_reintento(
                lambda: svc.users().messages().send(userId="me", body={"raw": raw}).execute()
            )
            logging.info(f"Correo HTML enviado. ID: {result['id']}")
            return True

        except Exception as e:
            logging.error(f"Error enviando correo HTML: {e}")
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
            svc = self.gmail_service
            result = _enviar_con_reintento(
                lambda: svc.users().messages().send(userId='me', body=send_message).execute()
            )
            logging.info(f"Correo enviado exitosamente. ID: {result['id']}")
            return True
            
        except Exception as e:
            logging.error(f"Error enviando correo: {e}")
            return False