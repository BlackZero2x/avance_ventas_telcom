"""
Proceso GUILLERMO — solo WhatsApp:
El correo a Guillermo ahora va consolidado en JefesEmailProcess junto con Carlos y Jesús.
"""
import logging
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "shared"))
from msg_utils import pick_variant
from execution_log import registrar_canal, canal_ok
from avance_finder import buscar_avance


class GuillermnoProcess:
    def __init__(self, config, wa, gmail_service=None):
        self.config = config
        self.wa = wa

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO GUILLERMO (WhatsApp)")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        ok_wa = self._enviar_whatsapp(archivo)
        logging.info("PROCESO GUILLERMO " + ("COMPLETADO" if ok_wa else "CON ERRORES"))
        return ok_wa

    def _enviar_whatsapp(self, archivo):
        if canal_ok("guillermo", "wa"):
            logging.info("WhatsApp Guillermo ya enviado hoy — omitiendo reenvio")
            return True
        caption = pick_variant(
            self.config.get("guillermo_message_variants"),
            self.config["guillermo_message"]
        )
        logging.info(f"Enviando archivo a Guillermo (WA): {os.path.basename(archivo)}")
        result = self.wa.send_file(self.config["guillermo_wa_contact"], archivo, caption=caption)
        # Verificar que wa_client retornó éxito (no solo que no tuvo excepción)
        if result.get("success"):
            registrar_canal("guillermo", "wa", True)
            return True
        else:
            logging.error(f"Error enviando a Guillermo por WA: {result.get('error', 'unknown')}")
            registrar_canal("guillermo", "wa", False)
            return False

    def _buscar_avance(self):
        return buscar_avance(self.config["archivos_avance_dir"], self.config)
