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
  5. GuillermnoProcess — envia AVANCE_{fecha}.xlsx a Guillermo por WhatsApp + correo (Gmail)
  6. CarlosProcess  — envia AVANCE_{fecha}.xlsx a Carlos por correo (Gmail)
  6. ItaloProcess   — carga AVANCE_VTAS_APPVENTORY → Sheets → notifica Italo
"""
import json
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

# Cargar .env antes que todo lo demás
def _cargar_dotenv():
    env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip())

_cargar_dotenv()

def _configurar_proxy():
    # Google APIs necesitan conexión directa — nunca van por proxy.
    # Para el resto de tráfico usamos el proxy corporativo con credenciales.
    no_proxy = "localhost,127.0.0.1,.googleapis.com,.google.com,.gstatic.com"
    os.environ["NO_PROXY"]    = no_proxy
    os.environ["no_proxy"]    = no_proxy

    proxy_host = os.environ.get("HTTP_PROXY", "")
    user       = os.environ.get("PROXY_USER", "")
    password   = os.environ.get("PROXY_PASS", "")

    if proxy_host and user and password:
        # Construir URL con credenciales: http://user:pass@host:port
        # El dominio AD puede contener '\' — codificarlo para la URL
        encoded_user = quote(user, safe="")
        encoded_pass = quote(password, safe="")
        base = proxy_host.replace("http://", "").replace("https://", "")
        proxy_url = f"http://{encoded_user}:{encoded_pass}@{base}"
    elif proxy_host:
        proxy_url = proxy_host
    else:
        proxy_url = ""

    if proxy_url:
        os.environ["HTTP_PROXY"]  = proxy_url
        os.environ["http_proxy"]  = proxy_url
    else:
        for var in ("HTTP_PROXY", "http_proxy"):
            os.environ.pop(var, None)

    # Google auth lib usa requests, que respeta HTTPS_PROXY para oauth2.
    # Forzar que nunca use proxy para HTTPS (las credenciales van por TLS directo).
    for var in ("HTTPS_PROXY", "https_proxy"):
        os.environ.pop(var, None)

_configurar_proxy()

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
from supervisor_alert_process import SupervisorAlertProcess
from shared.execution_log import (
    registrar_trigger, registrar_avance_ok, registrar_avance_fallo,
    registrar_modulo, registrar_fin,
)

# ── Logging ────────────────────────────────────────────────────────────────────
_avance_dir = os.environ.get("AVANCE_DIR", "C:/proyectos/AVANCE_MOVISTAR")
log_dir = Path(_avance_dir) / "logs"
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

CONFIG_PATH    = os.environ.get("CONFIG_PATH", f"{_avance_dir}/config.json")
CHECK_INTERVAL = int(os.environ.get("CHECK_INTERVAL_MINUTES", "10")) * 60


class Orchestrator:
    def __init__(self):
        logging.info("=" * 70)
        logging.info("ORQUESTADOR v2 - INICIO")
        logging.info("=" * 70)

        self.config = self._cargar_config(CONFIG_PATH)
        self._validar_env_emails()
        os.makedirs(self.config["temp_dir"], exist_ok=True)
        os.makedirs(self.config["archivos_avance_dir"], exist_ok=True)

        self.creds          = None
        self.gmail_service  = None
        self.sheets_service = None
        self.wa             = None
        self.processed_ids  = self._cargar_procesados()

    @staticmethod
    def _cargar_config(path: str) -> dict:
        claves_requeridas = [
            "email_sender", "email_subject", "processed_emails_file",
            "temp_dir", "archivos_avance_dir",
            "google_credentials_path", "google_token_path",
        ]
        if not os.path.exists(path):
            logging.error(f"No se encontró config.json en: {path}")
            sys.exit(1)
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except json.JSONDecodeError as e:
            logging.error(f"config.json tiene formato JSON inválido: {e}")
            sys.exit(1)
        faltantes = [k for k in claves_requeridas if k not in cfg]
        if faltantes:
            logging.error(f"config.json incompleto. Claves faltantes: {faltantes}")
            sys.exit(1)
        return cfg

    @staticmethod
    def _validar_env_emails():
        faltantes = [v for v in ("GUILLERMO_EMAIL", "CARLOS_EMAIL") if not os.environ.get(v, "").strip()]
        if faltantes:
            logging.error(f"Variables de entorno requeridas no definidas en .env: {faltantes}")
            sys.exit(1)

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

    @staticmethod
    def _notificar_inicio_avance():
        """Muestra un toast de Windows y espera 15s para que el usuario cierre Excel."""
        mensaje = (
            "El proceso AVANCE MOVISTAR se ejecutara en 15 segundos.\n"
            "Por favor guarda y cierra Excel para evitar conflictos."
        )
        try:
            # Toast nativo via PowerShell (Windows 10/11, no requiere librerías extra)
            ps_script = (
                "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null;"
                "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType=WindowsRuntime] | Out-Null;"
                "$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent("
                "    [Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
                "$template.SelectSingleNode('//text[@id=1]').InnerText = 'AVANCE MOVISTAR';"
                f"$template.SelectSingleNode('//text[@id=2]').InnerText = '{mensaje}';"
                "$toast = [Windows.UI.Notifications.ToastNotification]::new($template);"
                "$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('AVANCE MOVISTAR');"
                "$notifier.Show($toast);"
            )
            import subprocess as _sp
            _sp.Popen(
                ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_script],
                creationflags=0x08000000,  # CREATE_NO_WINDOW
            )
        except Exception as e:
            logging.warning(f"  No se pudo mostrar notificacion toast: {e}")

        # Adicionalmente: popup de MessageBox no bloqueante via PowerShell
        try:
            popup_script = (
                "Add-Type -AssemblyName System.Windows.Forms;"
                "[System.Windows.Forms.MessageBox]::Show("
                "    'El proceso AVANCE MOVISTAR iniciara en 15 segundos.`nCierra Excel para evitar conflictos.',"
                "    'AVANCE MOVISTAR',"
                "    [System.Windows.Forms.MessageBoxButtons]::OK,"
                "    [System.Windows.Forms.MessageBoxIcon]::Warning"
                ") | Out-Null"
            )
            import subprocess as _sp
            _sp.Popen(
                ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", popup_script],
                creationflags=0x08000000,
            )
        except Exception as e:
            logging.warning(f"  No se pudo mostrar popup: {e}")

        logging.info("  Esperando 15s para que el usuario cierre Excel...")
        time.sleep(15)

        # Matar cualquier instancia de Excel que siga abierta tras el aviso.
        try:
            import subprocess as _sp
            result = _sp.run(
                ["taskkill", "/F", "/IM", "excel.exe"],
                capture_output=True, text=True,
            )
            if "excel.exe" in result.stdout.lower():
                logging.info("  Excel cerrado forzosamente antes de ejecutar AVANCE.py")
            else:
                logging.info("  Sin instancias de Excel activas al iniciar.")
        except Exception as _e:
            logging.warning(f"  No se pudo cerrar Excel: {_e}")

    def _ejecutar_avance(self):
        self._notificar_inicio_avance()
        logging.info("[0/6] Ejecutando AVANCE.py para generar archivos del dia...")
        import subprocess
        script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AVANCE.py")
        periodo = self.config.get("periodo", "")
        cmd = [sys.executable, script]
        if periodo:
            cmd.append(periodo)
        result = subprocess.run(
            cmd,
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

    def _avisar_trigger_llego_si_corresponde(self):
        """Si el watchdog habia avisado a jefes que no llego info, notifica que ya llego."""
        from shared.execution_log import estado_hoy
        if not estado_hoy().get("jose_avisado", False):
            return
        try:
            sys.path.insert(0, str(self.config.get("whatsapp_server_dir",
                            str(Path(__file__).parent / "whatsapp_server"))))
            from msg_utils import pick_variant
            grupo_jefes = self.config.get("supervisor_alert_jefes_group", "")
            if grupo_jefes and self.wa.is_ready():
                variants = self.config.get("message_variants", {})
                msg = pick_variant(
                    variants.get("jefes_trigger_llego", []),
                    "Hola a todos, el avance de ventas ya esta disponible. Disculpen la demora, se envia a continuacion.",
                )
                self.wa.send_text(grupo_jefes, msg)
                logging.info("[OK] Aviso 'trigger llego' enviado al grupo jefes")
        except Exception as e:
            logging.warning(f"No se pudo enviar aviso 'trigger llego': {e}")

    def _run_all(self, email_id):
        logging.info("=" * 70)
        logging.info("EJECUTANDO TODOS LOS PROCESOS")
        logging.info("=" * 70)

        self._avisar_trigger_llego_si_corresponde()

        if not self._ejecutar_avance():
            self._marcar_procesado(email_id)
            registrar_avance_fallo("AVANCE.py termino con error")
            registrar_fin(False, "AVANCE.py termino con error")
            return

        registrar_avance_ok()

        if not self.wa.is_ready():
            logging.warning("WhatsApp no disponible antes de iniciar procesos. Esperando...")
            self._reconectar_whatsapp()

        pasos = [
            ("backs",            lambda: BacksProcess(self.config, self.sheets_service, self.wa).execute()),
            ("jefes",            lambda: JefesProcess(self.config, self.wa).execute()),
            ("jesus",            lambda: JesusProcess(self.config, self.wa).execute()),
            ("cristian",         lambda: CristianProcess(self.config, self.wa).execute()),
            ("guillermo",        lambda: GuillermnoProcess(self.config, self.wa, self.gmail_service).execute()),
            ("carlos",           lambda: CarlosProcess(self.config, self.gmail_service, self.wa).execute()),
            ("italo",            lambda: ItaloProcess(self.config, self.sheets_service, self.wa).execute()),
            ("supervisor_alert", lambda: SupervisorAlertProcess(self.config, self.wa).execute()),
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
            registrar_modulo(nombre, bool(results[nombre]))
            time.sleep(3)

        self._marcar_procesado(email_id)

        ok    = sum(1 for v in results.values() if v)
        total = len(results)
        todos_ok = ok == total
        registrar_fin(todos_ok, None if todos_ok else f"{total - ok} modulo(s) fallaron")

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
                registrar_trigger(email_id)
                self._run_all(email_id)
                logging.info("Orquestador finalizado exitosamente.")
                sys.exit(0)

            logging.info(f"  Sin trigger. Siguiente revision en {CHECK_INTERVAL // 60} minutos.")
            time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    Orchestrator().start()
