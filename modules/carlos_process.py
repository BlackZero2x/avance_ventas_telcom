"""
Proceso CARLOS:
1. Buscar AVANCE_{aaaa_mm_dd}.xlsx en Archivos_Avance (fecha de ayer)
2. Enviar texto de saludo a Carlos (abre el chat si no existe)
3. Enviar el archivo Excel como adjunto
"""
import logging
import os
import time
from datetime import datetime, timedelta
from msg_utils import pick_variant


class CarlosProcess:
    def __init__(self, config, wa):
        self.config = config
        self.wa = wa

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO CARLOS")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        contacto = self.config["carlos_wa_contact"]
        saludo = pick_variant(
            self.config.get("carlos_message_variants"),
            self.config["carlos_message"]
        )

        # open-wa requiere sendText antes de send_file en chats nuevos
        logging.info(f"Enviando saludo a Carlos: {saludo}")
        r1 = self.wa.send_text(contacto, saludo)
        logging.info(f"Resultado saludo: {r1}")
        time.sleep(5)

        logging.info(f"Enviando archivo a Carlos: {os.path.basename(archivo)}")
        r2 = self.wa.send_file(contacto, archivo, caption="")
        logging.info(f"Resultado archivo: {r2}")

        logging.info("PROCESO CARLOS COMPLETADO")
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
