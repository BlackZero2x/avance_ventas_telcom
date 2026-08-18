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
  5. GuillermnoProcess — envia AVANCE_{fecha}.xlsx a Guillermo por WhatsApp
  6. JefesEmailProcess — correo consolidado a Carlos + Guillermo + Jesús con captura TDS2 en el cuerpo
"""
import json
import logging
import os
import sys
import time
from datetime import datetime, date
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
from jefes_email_process import JefesEmailProcess
from italo_process import ItaloProcess
from supervisor_alert_process import SupervisorAlertProcess
from cuadro_resumen_sup_process import procesar_cuadro_resumen_sup
from cuadro_resumen_sup_diario_process import procesar_cuadro_resumen_sup_diario
from shared.execution_log import (
    registrar_trigger, registrar_avance_ok, registrar_avance_fallo,
    registrar_modulo, registrar_fin,
)
from shared.screenshot_safe import ScreenshotManager

# ── Verificación de datos SQL actualizados ─────────────────────────────────────

def _verificar_datos_frescos(periodo: str) -> tuple[bool, str]:
    """
    Verifica que en SQL Server haya registros de ALTAS con fecha_alta = ayer (D-1).
    Retorna (ok: bool, mensaje: str).
    """
    import urllib.parse
    from datetime import date, timedelta

    try:
        from sqlalchemy import create_engine, text as sa_text

        server   = os.environ.get("SQL_SERVER",   r"AUREN22\AUREN")
        database = os.environ.get("SQL_DATABASE", "eAuren")
        user     = os.environ.get("SQL_USER",     "")
        password = os.environ.get("SQL_PASSWORD", "")

        params = urllib.parse.quote_plus(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={user};"
            f"PWD={password};"
        )
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

        ayer = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        anio_mes = periodo  # formato YYYY-MM

        query = sa_text("""
            SELECT COUNT(*) AS cnt,
                   CONVERT(CHAR(10), MAX(CAST(a.fecha_alta AS DATE)), 126) AS max_fecha
            FROM fija_altas a
            WHERE FORMAT(a.fecha_alta, 'yyyy-MM') = :periodo
        """)

        with engine.connect() as conn:
            row = conn.execute(query, {"periodo": anio_mes}).fetchone()

        cnt       = row[0] if row else 0
        max_fecha = row[1] if row else None

        if cnt == 0:
            return False, f"SQL no tiene registros de ALTAS para el periodo {anio_mes}"

        if max_fecha != ayer:
            return False, (
                f"Datos SQL desactualizados: max fecha_alta = {max_fecha}, "
                f"se esperaba {ayer} (D-1)"
            )

        return True, f"Datos OK: {cnt} ALTAS, max fecha = {max_fecha}"

    except Exception as e:
        return False, f"Error al consultar SQL para verificar datos: {e}"

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

# Feriados en Perú (día, mes)
FERIADOS_PERU = {
    (28, 7),  # 28 de julio - Independencia del Perú
    (29, 7),  # 29 de julio - Feriado cívico / Festividad
}

def _es_feriado():
    """Retorna True si hoy es feriado en Perú."""
    hoy = date.today()
    return (hoy.day, hoy.month) in FERIADOS_PERU


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
            with open(path, "r", encoding="utf-8-sig") as f:
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
        faltantes = [v for v in ("CARLOS_EMAIL", "GUILLERMO_EMAIL", "JESUS_EMAIL") if not os.environ.get(v, "").strip()]
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

    def _obtener_ultimo_avance(self) -> str:
        """Obtiene la ruta del archivo AVANCE más reciente."""
        avance_dir = self.config.get("archivos_avance_dir", ".")
        archivos = sorted(
            (f for f in os.listdir(avance_dir) if f.startswith("AVANCE_") and f.endswith(".xlsx")),
            key=lambda x: os.path.getmtime(os.path.join(avance_dir, x)),
            reverse=True
        )
        if archivos:
            return os.path.join(avance_dir, archivos[0])
        return None

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
        # IMPORTANTE: se adquiere el lock global EXCEL_COM antes del taskkill,
        # porque un taskkill /IM excel.exe mata TODAS las instancias del sistema,
        # incluida la de otro proyecto (p.ej. SSFF) que en ese momento tenga Excel
        # abierto legítimamente bajo el mismo lock. Sin esta serialización, el
        # taskkill puede noquear la instancia COM de otro proceso a mitad de una
        # captura (visto en incidente SSFF_Corte_10AM del 18/08/2026).
        lock_taskkill = ScreenshotManager("MOVISTAR_TASKKILL_PRE_AVANCE")
        if not lock_taskkill.adquirir_lock(timeout=120):
            logging.warning("  No se pudo adquirir lock EXCEL_COM para el taskkill — se omite el cierre forzado.")
        else:
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
            finally:
                lock_taskkill.liberar_lock()

    # Destinatarios del correo de diagnóstico cuando AVANCE.py falla.
    # Leer desde variable de entorno (lista separada por comas) o usar lista fija.
    _AVANCE_ERROR_RECIPIENTS = [
        r for r in os.environ.get(
            "AVANCE_ERROR_RECIPIENTS",
            "augusto.moreno@auren.com.pe,jose.huiza@auren.com.pe,"
            "sistemas@auren.com.pe",
        ).split(",") if r.strip()
    ]

    def _ejecutar_avance(self):
        self._notificar_inicio_avance()
        self._descargar_adjunto_integratel()
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
            stderr_texto = result.stderr or ""
            if stderr_texto:
                logging.error(stderr_texto[-2000:])
            logging.error("AVANCE.py termino con error — abortando")
            self._enviar_diagnostico_error(stderr_texto, result.stdout or "")
            return False
        logging.info("[OK] AVANCE.py completado")
        return True

    def _descargar_adjunto_integratel(self):
        """Descarga el Excel de riesgo de Integratel antes de correr AVANCE.py.
        Si falla o no hay correo nuevo, se mantiene el último archivo descargado
        (el pipeline siempre debe tener un archivo de riesgo disponible)."""
        try:
            from shared.integratel_helper import descargar_adjunto_integratel
            destino = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "integratel_riesgo.xlsx"
            )
            descargar_adjunto_integratel(self.gmail_service, destino)
        except Exception as e:
            logging.warning(f"[Integratel] No se pudo descargar el adjunto: {e}")

    def _enviar_diagnostico_error(self, stderr: str, stdout: str):
        """Envía correo de diagnóstico cuando AVANCE.py falla."""
        try:
            from shared.error_diagnostico import diagnosticar, formatear_html
            from shared.gmail_helper import GmailHelper

            diag     = diagnosticar(stderr, stdout)
            proyecto = os.path.dirname(os.path.abspath(__file__))
            html     = formatear_html(diag, stderr, str(log_file), proyecto,
                                      hostname="developer7")

            asunto = f"[AVANCE MOVISTAR] Error en AVANCE.py — {diag['titulo']}"
            destinatarios = self._AVANCE_ERROR_RECIPIENTS

            gmail = GmailHelper(self.gmail_service)
            ok = gmail.send_email_html(destinatarios, asunto, html)
            if ok:
                logging.info(f"[OK] Correo de diagnostico enviado a: {', '.join(destinatarios)}")
            else:
                logging.warning("No se pudo enviar el correo de diagnostico")
        except Exception as e:
            logging.error(f"Error al enviar correo de diagnostico: {e}")

    def _ejecutar_carga_aqp(self):
        import subprocess
        script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "carga_info_aqp", "carga_datos_hoja_aqp.py")
        logging.info("[AQP] Iniciando carga de datos Arequipa → Google Sheets...")
        try:
            result = subprocess.run(
                [sys.executable, script],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=300,
            )
            if result.stdout:
                for line in result.stdout.strip().splitlines():
                    logging.info("  [AQP] %s", line)
            if result.returncode == 0:
                logging.info("[AQP] carga_datos_hoja_aqp completado OK.")
            else:
                if result.stderr:
                    logging.error("[AQP] stderr: %s", result.stderr[-2000:])
                logging.warning("[AQP] carga_datos_hoja_aqp terminó con código %s.", result.returncode)
        except Exception as e:
            logging.warning("[AQP] carga_datos_hoja_aqp falló: %s", e)

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

    def _avisar_datos_desactualizados(self, motivo: str):
        """Envía WA al owner avisando que los datos SQL no están al día."""
        owner_id = os.environ.get("OWNER_WA_ID", "").strip()
        if not owner_id:
            logging.warning("OWNER_WA_ID no definido en .env — no se puede notificar al owner")
            return
        try:
            msg = (
                f"[AVANCE MOVISTAR] El trigger llego pero los datos SQL no estan actualizados.\n"
                f"Motivo: {motivo}\n"
                f"El proceso NO se ejecuto. Revisar base de datos."
            )
            self.wa.send_text(owner_id, msg)
            logging.info(f"[OK] Aviso de datos desactualizados enviado a owner ({owner_id})")
        except Exception as e:
            logging.error(f"Error al enviar aviso al owner: {e}")

    def _run_all(self, email_id):
        logging.info("=" * 70)
        logging.info("EJECUTANDO TODOS LOS PROCESOS")
        logging.info("=" * 70)

        self._avisar_trigger_llego_si_corresponde()

        # ── Verificar que los datos SQL estén al día (D-1) ────────────────────
        periodo = self.config.get("periodo", "")
        if periodo:
            logging.info(f"Verificando datos SQL para periodo {periodo}...")
            datos_ok, motivo = _verificar_datos_frescos(periodo)
            if datos_ok:
                logging.info(f"[OK] {motivo}")
            else:
                logging.error(f"[DATOS DESACTUALIZADOS] {motivo}")
                self._avisar_datos_desactualizados(motivo)
                self._marcar_procesado(email_id)
                registrar_avance_fallo(motivo)
                registrar_fin(False, motivo)
                return
        else:
            logging.warning("Periodo no definido en config — se omite verificacion de datos")

        if not self._ejecutar_avance():
            self._marcar_procesado(email_id)
            registrar_avance_fallo("AVANCE.py termino con error")
            registrar_fin(False, "AVANCE.py termino con error")
            return

        registrar_avance_ok()

        # AQP en standby: pendiente actualizar SHEET_ID antes de reactivar.
        # self._ejecutar_carga_aqp()

        if not self.wa.is_ready():
            logging.warning("WhatsApp no disponible antes de iniciar procesos. Esperando...")
            self._reconectar_whatsapp()

        temp_dir = self.config.get("temp_dir", "")
        captura_tds1 = os.path.join(temp_dir, "captura_tds_1.png")
        captura_tds2 = os.path.join(temp_dir, "captura_tds_2.png")

        pasos = [
            ("backs",            lambda: BacksProcess(self.config, self.sheets_service, self.wa).execute()),
            ("jefes",            lambda: JefesProcess(self.config, self.wa).execute()),
            ("jesus",            lambda: JesusProcess(self.config, self.wa).execute()),
            ("cristian",         lambda: CristianProcess(self.config, self.wa).execute()),
            ("guillermo",        lambda: GuillermnoProcess(self.config, self.wa, self.gmail_service).execute()),
            ("carlos",           lambda: CarlosProcess(self.config, self.gmail_service, self.wa).execute()),
            ("jefes_email",      lambda: JefesEmailProcess(self.config, self.gmail_service).execute(
                                     captura_tds1=captura_tds1 if os.path.exists(captura_tds1) else None,
                                     captura_tds2=captura_tds2 if os.path.exists(captura_tds2) else None,
                                 )),
            ("italo",            lambda: ItaloProcess(self.config, self.sheets_service, self.wa).execute()),
            ("supervisor_alert", lambda: SupervisorAlertProcess(self.config, self.wa).execute()),
            ("cuadro_resumen_sup", lambda: procesar_cuadro_resumen_sup(self._obtener_ultimo_avance())),
            ("cuadro_resumen_sup_diario", lambda: procesar_cuadro_resumen_sup_diario(self._obtener_ultimo_avance())),
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
        # Verificar si es feriado en Perú
        if _es_feriado():
            fecha_hoy = date.today().strftime("%d de julio de %Y")
            logging.info("=" * 70)
            logging.info(f"[FERIADO] Hoy es {fecha_hoy} — Proceso AVANCE MOVISTAR deshabilitado.")
            logging.info("=" * 70)
            sys.exit(0)

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
