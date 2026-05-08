"""
Proceso GUILLERMO — WhatsApp + correo electrónico:
1. Buscar AVANCE_{aaaa-mm-dd}.xlsx en Archivos_Avance (fecha de ayer)
2. Enviar el archivo por WhatsApp con mensaje personalizado
3. Enviar el archivo por correo al destinatario definido en config["guillermo_email"]
   Asunto: Avance de FIJA actualizado al DD-MM-YYYY
"""
import logging
import os
from datetime import datetime, timedelta

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "shared"))
from gmail_helper import GmailHelper
from msg_utils import pick_variant


class GuillermnoProcess:
    def __init__(self, config, wa, gmail_service):
        self.config = config
        self.wa = wa
        self.gmail = GmailHelper(gmail_service)

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO GUILLERMO (WhatsApp + email)")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        ok_wa = self._enviar_whatsapp(archivo)
        ok_mail = self._enviar_correo(archivo)

        if ok_wa and ok_mail:
            logging.info("PROCESO GUILLERMO COMPLETADO")
        else:
            logging.warning(
                f"PROCESO GUILLERMO con errores — WA: {'OK' if ok_wa else 'FALLO'}, "
                f"email: {'OK' if ok_mail else 'FALLO'}"
            )
        return ok_wa and ok_mail

    def _enviar_whatsapp(self, archivo):
        caption = pick_variant(
            self.config.get("guillermo_message_variants"),
            self.config["guillermo_message"]
        )
        logging.info(f"Enviando archivo a Guillermo (WA): {os.path.basename(archivo)}")
        self.wa.send_file(self.config["guillermo_wa_contact"], archivo, caption=caption)
        return True

    def _enviar_correo(self, archivo):
        ayer = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
        asunto = f"Avance de FIJA actualizado al {ayer}"
        cuerpo = (
            f"Buen dia,\n\n"
            f"Adjunto el reporte de avance de ventas FIJA actualizado al {ayer}.\n\n"
            f"Saludos."
        )
        destinatario = self.config.get("guillermo_email", "")
        logging.info(f"Enviando correo a: {destinatario}")
        return self.gmail.send_email_with_attachment(
            [destinatario], asunto, cuerpo, archivo
        )

    def _buscar_avance(self):
        directorio = self.config["archivos_avance_dir"]
        ayer = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        nombre_esperado = os.path.join(directorio, f"AVANCE_{ayer}.xlsx")

        if os.path.exists(nombre_esperado):
            logging.info(f"Archivo encontrado: {nombre_esperado}")
            return nombre_esperado

        logging.error(
            f"No se encontro AVANCE_{ayer}.xlsx en {directorio}. "
            f"Ejecuta AVANCE.py primero para generar el archivo del dia."
        )
        return None
