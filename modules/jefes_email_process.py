"""
Proceso JEFES EMAIL — correo consolidado a Guillermo y Jesús (Carlos solo por WhatsApp):
- Cuerpo HTML con las capturas TDS1 y TDS2 incrustadas
- Adjunto: AVANCE_{fecha}.xlsx
- Destinatarios leídos de variables de entorno: GUILLERMO_EMAIL, JESUS_EMAIL, JOSE_EMAIL
- Carlos recibe solo por WhatsApp desde CarlosProcess
"""
import logging
import os
from datetime import datetime, timedelta

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "shared"))
from gmail_helper import GmailHelper
from execution_log import registrar_canal, canal_ok
from avance_finder import buscar_avance
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))

_HTML_TEMPLATE = """\
<!DOCTYPE html>
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222;">
  <p>Buen d&iacute;a,</p>
  <p>Adjunto el reporte de avance de ventas FIJA actualizado al <strong>{fecha}</strong>.</p>
  {bloque_tds1}
  {bloque_tds2}
  <p>Saludos.</p>
</body>
</html>
"""

_BLOQUE_IMG = '<p><img src="cid:{cid}" style="max-width:100%;"></p>'


class JefesEmailProcess:
    def __init__(self, config, gmail_service):
        self.config = config
        self.gmail = GmailHelper(gmail_service)

    def execute(self, captura_tds1=None, captura_tds2=None):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO JEFES EMAIL")
        logging.info("=" * 50)

        archivo = self._buscar_avance()
        if not archivo:
            return False

        img1 = captura_tds1 or self._buscar_captura("captura_tds_1.png")
        img2 = captura_tds2 or self._buscar_captura("captura_tds_2.png")

        ok = self._enviar_correo(archivo, img1, img2)
        logging.info("PROCESO JEFES EMAIL " + ("COMPLETADO" if ok else "CON ERRORES"))
        return ok

    # ── helpers ────────────────────────────────────────────────────────────────

    def _buscar_avance(self):
        return buscar_avance(self.config["archivos_avance_dir"], self.config)

    def _buscar_captura(self, nombre):
        ruta = os.path.join(self.config.get("temp_dir", ""), nombre)
        if os.path.exists(ruta):
            return ruta
        logging.warning(f"  {nombre} no encontrada en temp_dir")
        return None

    def _enviar_correo(self, archivo_avance, imagen_tds1, imagen_tds2):
        destinatarios = []
        # Excluir CARLOS_EMAIL del correo (solo por WA en CarlosProcess)
        for var in ("GUILLERMO_EMAIL", "JESUS_EMAIL", "JOSE_EMAIL"):
            valor = os.environ.get(var, "").strip()
            if valor:
                destinatarios.append(valor)
            else:
                logging.warning(f"  {var} no definido en .env — destinatario omitido")

        if not destinatarios:
            logging.error("  Ningun destinatario de email configurado — correo no enviado")
            return False

        if canal_ok("jefes_email", "email"):
            logging.info("  Correo jefes ya enviado hoy — omitiendo reenvio")
            return True

        ayer = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
        asunto = f"Avance de FIJA actualizado al {ayer}"

        bloque1 = _BLOQUE_IMG.format(cid="tds_image_1") if imagen_tds1 else ""
        bloque2 = _BLOQUE_IMG.format(cid="tds_image_2") if imagen_tds2 else ""
        cuerpo_html = _HTML_TEMPLATE.format(fecha=ayer, bloque_tds1=bloque1, bloque_tds2=bloque2)

        logging.info(f"  Enviando correo a: {destinatarios}")
        logging.info(f"  Adjunto: {os.path.basename(archivo_avance)}")

        ok = self.gmail.send_email_with_two_images(
            recipients=destinatarios,
            subject=asunto,
            html_body=cuerpo_html,
            image_path_1=imagen_tds1,
            image_path_2=imagen_tds2,
            attachment_path=archivo_avance,
        )
        if ok:
            registrar_canal("jefes_email", "email", True)
            logging.info("  Correo jefes enviado OK")
        else:
            logging.error("  Error enviando correo jefes")
        return ok
