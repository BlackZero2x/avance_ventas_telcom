"""
Proceso CARLOS — envío por correo electrónico:
1. Buscar AVANCE_{aaaa-mm-dd}.xlsx en Archivos_Avance (fecha de ayer)
2. Enviar el archivo por Gmail al destinatario definido en config["carlos_email"]
   Asunto: Avance de FIJA actualizado al DD-MM-YYYY
"""
import logging
import os
from datetime import datetime, timedelta

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "shared"))
from gmail_helper import GmailHelper


class CarlosProcess:
    def __init__(self, config, gmail_service):
        self.config = config
        self.gmail = GmailHelper(gmail_service)

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO CARLOS (email)")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        ayer = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
        asunto = f"Avance de FIJA actualizado al {ayer}"
        cuerpo = (
            f"Buen dia,\n\n"
            f"Adjunto el reporte de avance de ventas FIJA actualizado al {ayer}.\n\n"
            f"Saludos."
        )

        destinatario = self.config.get("carlos_email", "")
        logging.info(f"Enviando correo a: {destinatario}")
        logging.info(f"Adjunto: {os.path.basename(archivo)}")
        ok = self.gmail.send_email_with_attachment(
            [destinatario], asunto, cuerpo, archivo
        )

        if ok:
            logging.info("PROCESO CARLOS COMPLETADO")
        else:
            logging.error("PROCESO CARLOS FALLO al enviar correo")
        return ok

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
