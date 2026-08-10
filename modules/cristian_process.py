"""
Proceso CRISTIAN:
1. Buscar AVANCE_{aaaa_mm_dd}.xlsx en Archivos_Avance (por defecto fecha de ayer)
2. Enviar el archivo a Cristian por WhatsApp con mensaje personalizado
"""
import logging
import os
from msg_utils import pick_variant
from shared.avance_finder import buscar_avance


class CristianProcess:
    def __init__(self, config, wa):
        self.config = config
        self.wa = wa

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO CRISTIAN")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        logging.info(f"Enviando archivo a Cristian: {os.path.basename(archivo)}")
        caption = pick_variant(self.config.get("cristian_message_variants"), self.config["cristian_message"])
        result = self.wa.send_file(self.config["cristian_wa_contact"], archivo, caption=caption)

        if result.get("success"):
            logging.info("PROCESO CRISTIAN COMPLETADO")
            return True
        else:
            logging.error(f"Error enviando a Cristian: {result.get('error', 'unknown')}")
            return False

    def _buscar_avance(self):
        return buscar_avance(self.config["archivos_avance_dir"], self.config)
