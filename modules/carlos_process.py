"""
Proceso CARLOS — WhatsApp + correo electrónico:
1. Buscar AVANCE_{aaaa-mm-dd}.xlsx en Archivos_Avance (fecha de ayer)
2. Enviar el archivo por WhatsApp con mensaje personalizado
3. Enviar el archivo SEGUIMIENTO_VDD_FIJA por WhatsApp
4. Enviar el archivo AVANCE por Gmail al destinatario definido en CARLOS_EMAIL
"""
import glob
import logging
import os
from datetime import datetime, timedelta

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "shared"))
from gmail_helper import GmailHelper
from msg_utils import pick_variant
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))


class CarlosProcess:
    def __init__(self, config, gmail_service, wa=None):
        self.config = config
        self.gmail = GmailHelper(gmail_service)
        self.wa = wa

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO CARLOS (WhatsApp + email)")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        ok_wa   = self._enviar_whatsapp(archivo)
        ok_mail = self._enviar_correo(archivo)

        if ok_wa and ok_mail:
            logging.info("PROCESO CARLOS COMPLETADO")
        else:
            logging.warning(
                f"PROCESO CARLOS con errores — WA: {'OK' if ok_wa else 'FALLO'}, "
                f"email: {'OK' if ok_mail else 'FALLO'}"
            )
        return ok_mail  # el correo es el canal principal; WA es adicional

    def _enviar_whatsapp(self, archivo_avance):
        if not self.wa:
            logging.info("  WhatsApp no configurado para Carlos — omitiendo")
            return True

        contacto = self.config.get("carlos_wa_contact")
        if not contacto:
            logging.info("  carlos_wa_contact no definido — omitiendo WA")
            return True

        try:
            caption = pick_variant(
                self.config.get("carlos_message_variants"),
                self.config.get("carlos_message", "Buen dia Carlos, adjunto el avance actualizado.")
            )
            logging.info(f"  Enviando AVANCE a Carlos por WA: {os.path.basename(archivo_avance)}")
            self.wa.send_file(contacto, archivo_avance, caption=caption)

            self._enviar_seguimiento_wa(contacto)
            return True
        except Exception as e:
            logging.error(f"  Error enviando a Carlos por WA: {e}")
            return False

    def _enviar_seguimiento_wa(self, contacto):
        archivos_dir = self.config["archivos_avance_dir"]
        candidatos = glob.glob(os.path.join(archivos_dir, "SEGUIMIENTO_VDD_FIJA_*.xlsx"))
        if not candidatos:
            logging.warning("  No se encontro SEGUIMIENTO_VDD_FIJA para Carlos")
            return

        def _fecha_desde_nombre(path):
            nombre = os.path.basename(path)
            partes = nombre.replace("SEGUIMIENTO_VDD_FIJA_", "").replace(".xlsx", "").split("-")
            try:
                return datetime(int(partes[2]), int(partes[1]), int(partes[0]))
            except Exception:
                return datetime.min

        archivo_seg = max(candidatos, key=_fecha_desde_nombre)
        logging.info(f"  Enviando SEGUIMIENTO a Carlos: {os.path.basename(archivo_seg)}")
        self.wa.send_file(contacto, archivo_seg, caption="")

    def _enviar_correo(self, archivo):
        ayer = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
        asunto = f"Avance de FIJA actualizado al {ayer}"
        cuerpo = (
            f"Buen dia,\n\n"
            f"Adjunto el reporte de avance de ventas FIJA actualizado al {ayer}.\n\n"
            f"Saludos."
        )

        destinatario = os.environ.get("CARLOS_EMAIL", "").strip()
        if not destinatario:
            logging.error("CARLOS_EMAIL no definido en .env — correo no enviado")
            return False
        logging.info(f"  Enviando correo a: {destinatario}")
        logging.info(f"  Adjunto: {os.path.basename(archivo)}")
        ok = self.gmail.send_email_with_attachment(
            [destinatario], asunto, cuerpo, archivo
        )
        if ok:
            logging.info("  Correo Carlos enviado OK")
        else:
            logging.error("  Error enviando correo Carlos")
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
