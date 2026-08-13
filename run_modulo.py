"""
Ejecuta modulos individualmente o todos en modo REAL (destinatarios reales).
Incluye autenticacion Google y verificacion WhatsApp igual que main_v2.py,
pero sin esperar trigger de correo — ejecucion inmediata y manual.

Uso:
    python run_modulo.py backs
    python run_modulo.py jefes
    python run_modulo.py jesus
    python run_modulo.py cristian
    python run_modulo.py carlos/guillermo
    python run_modulo.py italo
    python run_modulo.py todos
"""
import json
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

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
    no_proxy = "localhost,127.0.0.1,.googleapis.com,.google.com,.gstatic.com"
    os.environ["NO_PROXY"] = no_proxy
    os.environ["no_proxy"] = no_proxy

    proxy_host = os.environ.get("HTTP_PROXY", "")
    user       = os.environ.get("PROXY_USER", "")
    password   = os.environ.get("PROXY_PASS", "")

    if proxy_host and user and password:
        encoded_user = quote(user, safe="")
        encoded_pass = quote(password, safe="")
        base = proxy_host.replace("http://", "").replace("https://", "")
        proxy_url = f"http://{encoded_user}:{encoded_pass}@{base}"
    elif proxy_host:
        proxy_url = proxy_host
    else:
        proxy_url = ""

    if proxy_url:
        os.environ["HTTP_PROXY"] = proxy_url
        os.environ["http_proxy"] = proxy_url
    else:
        for var in ("HTTP_PROXY", "http_proxy"):
            os.environ.pop(var, None)

    for var in ("HTTPS_PROXY", "https_proxy"):
        os.environ.pop(var, None)

_configurar_proxy()

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "modules"))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "whatsapp_server"))

from modules.shared.execution_log import registrar_modulo, registrar_avance_ok, registrar_avance_fallo, registrar_fin
from wa_client import WhatsAppClient
from backs_process import BacksProcess
from jefes_process import JefesProcess
from jefes_email_process import JefesEmailProcess
from jesus_process import JesusProcess
from cristian_process import CristianProcess
from guillermo_process import GuillermnoProcess
from carlos_process import CarlosProcess
from italo_process import ItaloProcess
from supervisor_alert_process import SupervisorAlertProcess

# ── Logging ────────────────────────────────────────────────────────────────────
log_dir = Path("C:/proyectos/AVANCE_MOVISTAR/logs")
log_dir.mkdir(exist_ok=True)
log_file = log_dir / f"manual_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

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

CONFIG_PATH = "C:/proyectos/AVANCE_MOVISTAR/config.json"

with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
    config = json.load(f)

os.makedirs(config["temp_dir"], exist_ok=True)
os.makedirs(config["archivos_avance_dir"], exist_ok=True)


# ── Autenticacion Google ───────────────────────────────────────────────────────

def _autenticar_google():
    logging.info("Autenticando con Google APIs...")
    token_path = config["google_token_path"]
    creds_path = config["google_credentials_path"]
    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())

    sheets_service = build("sheets", "v4", credentials=creds)
    gmail_service  = build("gmail",  "v1", credentials=creds)
    logging.info("[OK] Autenticacion Google exitosa")
    return sheets_service, gmail_service


# ── Conexion WhatsApp ──────────────────────────────────────────────────────────

def _conectar_whatsapp():
    wa = WhatsAppClient(
        host=config.get("wa_host", "localhost"),
        port=config.get("wa_port", 8002),
    )
    if not wa.is_ready():
        logging.error("Servidor WhatsApp no disponible. Asegurate de que wa_server.js este corriendo.")
        sys.exit(1)
    logging.info("[OK] WhatsApp listo")
    return wa


# ── AVANCE.py ─────────────────────────────────────────────────────────────────

def _ejecutar_avance(gmail_service=None):
    if gmail_service is not None:
        try:
            from modules.shared.integratel_helper import descargar_adjunto_integratel
            destino = os.path.join(os.path.dirname(os.path.abspath(__file__)), "integratel_riesgo.xlsx")
            descargar_adjunto_integratel(gmail_service, destino)
        except Exception as e:
            logging.warning(f"[Integratel] No se pudo descargar el adjunto: {e}")
    logging.info("[0/6] Ejecutando AVANCE.py para generar archivos del dia...")
    import subprocess
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AVANCE.py")
    periodo = config.get("periodo", "")
    cmd = [sys.executable, script]
    if periodo:
        cmd.append(periodo)
    result = subprocess.run(
        cmd,
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    if result.stdout:
        for line in result.stdout.strip().splitlines():
            logging.info("  [AVANCE] %s", line.replace("�", "?"))
    if result.returncode != 0:
        if result.stderr:
            logging.error(result.stderr[-2000:])
        logging.error("AVANCE.py termino con error — abortando")
        return False
    logging.info("[OK] AVANCE.py completado")
    return True


# ── Modulos ────────────────────────────────────────────────────────────────────

def _run(modulo, sheets_service, gmail_service, wa, forzar=False):
    temp_dir = config.get("temp_dir", "")
    captura_tds1 = os.path.join(temp_dir, "captura_tds_1.png")
    captura_tds2 = os.path.join(temp_dir, "captura_tds_2.png")

    MODULOS = {
        "backs":            lambda: BacksProcess(config, sheets_service, wa).execute(),
        "jefes":            lambda: JefesProcess(config, wa).execute(),
        "jefes_email":      lambda: JefesEmailProcess(config, gmail_service).execute(
                                captura_tds1=captura_tds1 if os.path.exists(captura_tds1) else None,
                                captura_tds2=captura_tds2 if os.path.exists(captura_tds2) else None,
                            ),
        "jesus":            lambda: JesusProcess(config, wa).execute(),
        "cristian":         lambda: CristianProcess(config, wa).execute(),
        "guillermo":        lambda: GuillermnoProcess(config, wa, gmail_service).execute(),
        "carlos":           lambda: CarlosProcess(config, gmail_service, wa).execute(),
        "italo":            lambda: ItaloProcess(config, sheets_service, wa).execute(),
        "supervisor_alert": lambda: SupervisorAlertProcess(config, wa).execute(),
    }

    if modulo == "todos":
        from main_v2 import _verificar_datos_frescos
        periodo = config.get("periodo", "")
        if forzar:
            logging.warning(f"--forzar activo: se omite verificacion de datos frescos (periodo={periodo})")
        elif periodo:
            logging.info(f"Verificando datos SQL para periodo {periodo}...")
            datos_ok, motivo = _verificar_datos_frescos(periodo)
            if datos_ok:
                logging.info(f"[OK] {motivo}")
            else:
                logging.error(f"[DATOS DESACTUALIZADOS] {motivo}")
                logging.error("Abortando ejecucion manual. Usa --forzar para saltear esta verificacion.")
                sys.exit(1)
        else:
            logging.warning("Periodo no definido en config — se omite verificacion de datos")

        # Si ya existe el AVANCE del periodo, no regenerar
        from modules.shared.avance_finder import buscar_avance as _buscar_avance_existente
        archivo_existente = _buscar_avance_existente(config["archivos_avance_dir"], config)
        if archivo_existente:
            logging.info(f"Archivo AVANCE existente detectado: {archivo_existente}")
            logging.info("Saltando ejecucion de AVANCE.py (archivo ya generado)")
            registrar_avance_ok()
        else:
            avance_ok = _ejecutar_avance(gmail_service)
            if not avance_ok:
                registrar_avance_fallo("AVANCE.py termino con error")
                registrar_fin(False, "AVANCE.py termino con error")
                sys.exit(1)
            registrar_avance_ok()

        results = {}
        for nombre, fn in MODULOS.items():
            logging.info(f"--- Ejecutando: {nombre} ---")
            try:
                results[nombre] = fn()
            except Exception as e:
                logging.error(f"Error en {nombre}: {e}")
                import traceback
                logging.error(traceback.format_exc())
                results[nombre] = False
            registrar_modulo(nombre, bool(results[nombre]))
            time.sleep(3)

        ok = sum(1 for v in results.values() if v)
        todos_ok = (ok == len(results))
        logging.info("=" * 70)
        logging.info(f"RESUMEN: {ok}/{len(results)} procesos exitosos")
        for nombre, status in results.items():
            logging.info(f"  {'[OK]   ' if status else '[ERROR]'} {nombre}")
        logging.info("=" * 70)
        registrar_fin(todos_ok, None if todos_ok else f"{len(results) - ok} modulo(s) fallaron")

    elif modulo in MODULOS:
        try:
            result = MODULOS[modulo]()
            registrar_modulo(modulo, bool(result))
        except Exception as e:
            registrar_modulo(modulo, False)
            raise
    else:
        print(f"Modulo '{modulo}' no reconocido. Opciones: {', '.join(list(MODULOS) + ['todos'])}")
        sys.exit(1)


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python run_modulo.py <modulo> [--forzar]")
        print("Modulos disponibles: backs, jefes, jesus, cristian, guillermo, carlos, italo, supervisor_alert, todos")
        print("  --forzar  Salta la verificacion de datos frescos (util para periodos pasados)")
        sys.exit(1)

    modulo = sys.argv[1].lower()
    _forzar = "--forzar" in sys.argv

    sheets_service, gmail_service = _autenticar_google()
    wa = _conectar_whatsapp()

    _run(modulo, sheets_service, gmail_service, wa, forzar=_forzar)
