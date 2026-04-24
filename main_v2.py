"""
Orquestador v2 — ejecutado por el Programador de Tareas de Windows.

Comportamiento:
  - Arranca, autentica Google y verifica WhatsApp.
  - Revisa el correo trigger cada 10 minutos.
  - En cuanto detecta el trigger ejecuta todos los procesos y termina.
  - Si tras MAX_INTENTOS ciclos no llega el trigger, termina igual (sin colgar).

Trigger: correo de e@auren.com.pe con asunto "avance_ventas - Actualizacion disponible"

Procesos en orden:
  1. BacksProcess   — genera AVANCE_RESUMIDO → sube a Sheets → notifica grupo BACKS
  2. JefesProcess   — capturas TDS + SEGUIMIENTO_VDD → grupo JEFES
  3. JesusProcess   — envia AVANCE_{fecha}.xlsx a Jesus
  4. CristianProcess — envia AVANCE_{fecha}.xlsx a Cristian
  5. GuillermnoProcess — envia AVANCE_{fecha}.xlsx a Guillermo
  6. ItaloProcess   — carga AVANCE_VTAS_APPVENTORY → Sheets → notifica Italo
"""
import json
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path

# El proxy del sistema usa https:// en su URL pero solo soporta HTTP;
# corregimos ambas variables para que urllib3/requests funcionen con Google APIs.
os.environ["HTTPS_PROXY"] = "http://192.168.2.1:3128"
os.environ["HTTP_PROXY"]  = "http://192.168.2.1:3128"
os.environ["NO_PROXY"]    = "localhost,127.0.0.1"

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "modules"))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "whatsapp_server"))

from wa_client import WhatsAppClient
from backs_process import BacksProcess
from jefes_process import JefesProcess
from jesus_process import JesusProcess
from cristian_process import CristianProcess
from guillermo_process import GuillermnoProcess
from carlos_process import CarlosProcess
from italo_process import ItaloProcess

# ── Logging ────────────────────────────────────────────────────────────────────
log_dir = Path("C:/AVANCE_MOVISTAR/logs")
log_dir.mkdir(exist_ok=True)
log_file = log_dir / f"automation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

CONFIG_PATH    = "C:/AVANCE_MOVISTAR/config.json"
CHECK_INTERVAL = 10 * 60   # 10 minutos entre cada revision


class Orchestrator:
    def __init__(self):
        logging.info("=" * 70)
        logging.info("ORQUESTADOR v2 - INICIO")
        logging.info("=" * 70)

        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            self.config = json.load(f)

        os.makedirs(self.config["temp_dir"], exist_ok=True)
        os.makedirs(self.config["archivos_avance_dir"], exist_ok=True)

        self.creds          = None
        self.gmail_service  = None
        self.sheets_service = None
        self.wa             = None
        self.processed_ids  = self._cargar_procesados()

    # ── Email tracking ─────────────────────────────────────────────────────────

    def _cargar_procesados(self):
        path = self.config["processed_emails_file"]
        if os.path.exists(path):
            with open(path, "r") as f:
                return set(f.read().splitlines())
        return set()

    def _marcar_procesado(self, email_id):
        self.processed_ids.add(email_id)
        with open(self.config["processed_emails_file"], "a") as f:
            f.write(f"{email_id}\n")

    # ── Google Auth ────────────────────────────────────────────────────────────

    def _autenticar_google(self):
        logging.info("Autenticando con Google APIs...")
        token_path = self.config["google_token_path"]
        creds_path = self.config["google_credentials_path"]

        if os.path.exists(token_path):
            self.creds = Credentials.from_authorized_user_file(token_path, SCOPES)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
                self.creds = flow.run_local_server(port=0)
            with open(token_path, "w") as f:
                f.write(self.creds.to_json())

        self.gmail_service  = build("gmail",  "v1", credentials=self.creds)
        self.sheets_service = build("sheets", "v4", credentials=self.creds)
        logging.info("[OK] Autenticacion Google exitosa")

    # ── WhatsApp ───────────────────────────────────────────────────────────────

    def _conectar_whatsapp(self):
        self.wa = WhatsAppClient(
            host=self.config.get("wa_host", "localhost"),
            port=self.config.get("wa_port", 8002),
        )
        intentos = 0
        max_intentos = 60  # 60 x 10s = 10 minutos
        while not self.wa.is_ready():
            intentos += 1
            if intentos > max_intentos:
                logging.error("El servidor WhatsApp no respondio tras 10 minutos. Abortando.")
                return False
            logging.info(f"  WhatsApp no disponible. Reintento {intentos}/{max_intentos} en 10s...")
            time.sleep(10)
        logging.info("[OK] WhatsApp (open-wa) listo")
        return True

    # ── Email trigger ──────────────────────────────────────────────────────────

    def _buscar_trigger(self):
        try:
            query = (
                f'from:{self.config["email_sender"]} '
                f'subject:{self.config["email_subject"]} '
                f'is:unread'
            )
            results = self.gmail_service.users().messages().list(
                userId="me", q=query
            ).execute()

            for msg in results.get("messages", []):
                msg_id = msg["id"]
                if msg_id not in self.processed_ids:
                    logging.info(f"[OK] Trigger detectado (email id: {msg_id})")
                    self.gmail_service.users().messages().modify(
                        userId="me",
                        id=msg_id,
                        body={"removeLabelIds": ["UNREAD"]},
                    ).execute()
                    return msg_id

            return None

        except Exception as e:
            logging.error(f"Error verificando email: {e}")
            return None

    # ── Procesos ───────────────────────────────────────────────────────────────

    def _ejecutar_avance(self):
        logging.info("[0/6] Ejecutando AVANCE.py para generar archivos del dia...")
        import subprocess
        script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AVANCE.py")
        result = subprocess.run(
            [sys.executable, script],
            capture_output=True, text=True, encoding="utf-8", errors="replace"
        )
        if result.stdout:
            for line in result.stdout.strip().splitlines():
                logging.info(f"  [AVANCE] {line}")
        if result.returncode != 0:
            if result.stderr:
                logging.error(result.stderr[-2000:])
            logging.error("AVANCE.py termino con error — abortando")
            return False
        logging.info("[OK] AVANCE.py completado")
        return True

    def _reconectar_whatsapp(self):
        intentos = 0
        max_intentos = 60  # 60 x 10s = 10 minutos
        while not self.wa.is_ready():
            intentos += 1
            if intentos > max_intentos:
                logging.error("WhatsApp no respondio tras 10 minutos. Los procesos de envio fallaran.")
                return False
            logging.info(f"  WhatsApp caido. Esperando reconexion {intentos}/{max_intentos}...")
            time.sleep(10)
        logging.info("[OK] WhatsApp reconectado")
        return True

    def _run_all(self, email_id):
        logging.info("=" * 70)
        logging.info("EJECUTANDO TODOS LOS PROCESOS")
        logging.info("=" * 70)

        if not self._ejecutar_avance():
            self._marcar_procesado(email_id)
            return

        if not self.wa.is_ready():
            logging.warning("WhatsApp no disponible antes de iniciar procesos. Esperando...")
            self._reconectar_whatsapp()

        pasos = [
            ("backs",     lambda: BacksProcess(self.config, self.sheets_service, self.wa).execute()),
            ("jefes",     lambda: JefesProcess(self.config, self.wa).execute()),
            ("jesus",     lambda: JesusProcess(self.config, self.wa).execute()),
            ("cristian",  lambda: CristianProcess(self.config, self.wa).execute()),
            ("guillermo", lambda: GuillermnoProcess(self.config, self.wa).execute()),
            ("carlos",    lambda: CarlosProcess(self.config, self.wa).execute()),
            ("italo",     lambda: ItaloProcess(self.config, self.sheets_service, self.wa).execute()),
        ]

        results = {}
        for nombre, fn in pasos:
            try:
                results[nombre] = fn()
            except Exception as e:
                logging.error(f"Error en {nombre}: {e}")
                import traceback
                logging.error(traceback.format_exc())
                results[nombre] = False
            time.sleep(3)

        self._marcar_procesado(email_id)

        ok    = sum(1 for v in results.values() if v)
        total = len(results)
        logging.info("=" * 70)
        logging.info(f"RESUMEN: {ok}/{total} procesos exitosos")
        for nombre, status in results.items():
            logging.info(f"  {'[OK]   ' if status else '[ERROR]'} {nombre}")
        logging.info("=" * 70)

    # ── Loop principal ─────────────────────────────────────────────────────────

    def start(self):
        self._autenticar_google()

        if not self._conectar_whatsapp():
            logging.error("Abortando: WhatsApp no disponible")
            sys.exit(1)

        logging.info(f"Revisando correo cada {CHECK_INTERVAL // 60} minutos. Esperando trigger...")

        intento = 0
        while True:
            intento += 1
            logging.info(f"[Revision {intento}] Verificando trigger... ({datetime.now().strftime('%H:%M:%S')})")

            email_id = self._buscar_trigger()

            if email_id:
                self._run_all(email_id)
                logging.info("Orquestador finalizado exitosamente.")
                sys.exit(0)

            logging.info(f"  Sin trigger. Siguiente revision en {CHECK_INTERVAL // 60} minutos.")
            time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    Orchestrator().start()
