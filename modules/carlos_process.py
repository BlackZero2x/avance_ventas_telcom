"""
Proceso CARLOS — solo WhatsApp:
El correo a Carlos ahora va consolidado en JefesEmailProcess junto con Guillermo y Jesús.
"""
import glob
import logging
import os
from datetime import datetime

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "shared"))
from msg_utils import pick_variant
from execution_log import registrar_canal, canal_ok
from avance_finder import buscar_avance


class CarlosProcess:
    def __init__(self, config, gmail_service=None, wa=None):
        self.config = config
        self.wa = wa

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO CARLOS (WhatsApp)")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        ok_wa = self._enviar_whatsapp(archivo)
        logging.info("PROCESO CARLOS " + ("COMPLETADO" if ok_wa else "CON ERRORES"))
        return ok_wa

    def _enviar_whatsapp(self, archivo_avance):
        if not self.wa:
            logging.info("  WhatsApp no configurado para Carlos — omitiendo")
            return True

        contacto = self.config.get("carlos_wa_contact")
        if not contacto:
            logging.info("  carlos_wa_contact no definido — omitiendo WA")
            return True

        try:
            if canal_ok("carlos", "wa"):
                logging.info("  WhatsApp Carlos ya enviado hoy — omitiendo reenvio")
                return True
            caption = pick_variant(
                self.config.get("carlos_message_variants"),
                self.config.get("carlos_message", "Buen dia Carlos, adjunto el avance actualizado.")
            )
            logging.info(f"  Enviando AVANCE a Carlos por WA: {os.path.basename(archivo_avance)}")
            result = self.wa.send_file(contacto, archivo_avance, caption=caption)

            # Verificar que wa_client retornó éxito (no solo que no tuvo excepción)
            if not result.get("success"):
                logging.error(f"  Error enviando AVANCE a Carlos por WA: {result.get('error', 'unknown')}")
                registrar_canal("carlos", "wa", False)
                return False

            self._enviar_seguimiento_wa(contacto)
            registrar_canal("carlos", "wa", True)
            return True
        except Exception as e:
            logging.error(f"  Error enviando a Carlos por WA: {e}")
            registrar_canal("carlos", "wa", False)
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

    def _buscar_avance(self):
        return buscar_avance(self.config["archivos_avance_dir"], self.config)
