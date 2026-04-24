"""
Proceso GUILLERMO:
1. Buscar AVANCE_{aaaa_mm_dd}.xlsx en Archivos_Avance (por defecto fecha de ayer)
2. Enviar el archivo a Guillermo por WhatsApp con mensaje personalizado
"""
import logging
import os
from datetime import datetime, timedelta
from msg_utils import pick_variant


class GuillermnoProcess:
    def __init__(self, config, wa):
        self.config = config
        self.wa = wa

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO GUILLERMO")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        logging.info(f"Enviando archivo a Guillermo: {os.path.basename(archivo)}")
        caption = pick_variant(self.config.get("guillermo_message_variants"), self.config["guillermo_message"])
        self.wa.send_file(self.config["guillermo_wa_contact"], archivo, caption=caption)

        logging.info("PROCESO GUILLERMO COMPLETADO")
        return True

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
