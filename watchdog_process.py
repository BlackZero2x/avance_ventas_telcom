"""
watchdog_process.py
Ejecutado por el Programador de Tareas lun-vie a las:
  10:00, 11:00, 12:00 (manana) | 14:00, 16:00, 18:00 (tarde)

Logica de verificacion:
  1. Si el pipeline esta corriendo ahora mismo, no interferir.
  2. Si no hay trigger antes de las 13:00, solo loguear (es normal).
  3. Si no hay trigger a las 13:00+, avisar a jefes y a Jose Huiza (una sola vez).
  4. Si hay trigger: verificar el estado REAL del dia (archivos + modulos).
     - Si todo OK segun verificacion real: no hacer nada.
     - Si hay gaps: relanzar run_modulo.py con solo los modulos faltantes
       (o todos si AVANCE.py no genero archivos).
     - Limite: MAX_REINTENTOS intentos por dia; si se supera, solo loguear.
"""
import json
import logging
import os
import subprocess
import sys
import time
from datetime import datetime, date, timedelta
from pathlib import Path
from urllib.parse import quote

BASE_DIR = Path(__file__).parent

MAX_REINTENTOS = 3  # maximo de reintentos automaticos por dia

def _cargar_dotenv():
    env_path = BASE_DIR / ".env"
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

# ── Logging ────────────────────────────────────────────────────────────────────
log_dir = BASE_DIR / "logs"
log_dir.mkdir(exist_ok=True)
log_file = log_dir / f"watchdog_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

EXEC_LOG_PATH = BASE_DIR / "execution_log.json"
CONFIG_PATH   = BASE_DIR / "whatsapp_server" / "config.json"
HOY           = date.today().isoformat()  # "YYYY-MM-DD"


def _leer_estado_hoy() -> dict:
    if not EXEC_LOG_PATH.exists():
        return {}
    try:
        data = json.loads(EXEC_LOG_PATH.read_text(encoding="utf-8"))
        return data.get(HOY, {})
    except Exception as e:
        logging.error(f"No se pudo leer execution_log.json: {e}")
        return {}


def _actualizar_log(campos: dict):
    try:
        data = {}
        if EXEC_LOG_PATH.exists():
            data = json.loads(EXEC_LOG_PATH.read_text(encoding="utf-8"))
        if HOY not in data:
            data[HOY] = {}
        data[HOY].update(campos)
        EXEC_LOG_PATH.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        logging.error(f"No se pudo actualizar execution_log.json: {e}")


def _wa_client_y_config():
    """Devuelve (wa, config) listos para usar, o lanza excepcion si WA no esta disponible."""
    sys.path.insert(0, str(BASE_DIR / "whatsapp_server"))
    from wa_client import WhatsAppClient
    with open(CONFIG_PATH, encoding="utf-8-sig") as f:
        config = json.load(f)
    wa = WhatsAppClient(
        host=config.get("wa_host", "localhost"),
        port=config.get("wa_port", 8002),
    )
    if not wa.is_ready():
        raise RuntimeError("WhatsApp no disponible")
    return wa, config


def _avisar_sin_trigger() -> bool:
    try:
        from msg_utils import pick_variant
        wa, config = _wa_client_y_config()
        variants = config.get("message_variants", {})
        hora_lima = datetime.now().strftime("%I:%M %p").lstrip("0")
        grupo_jefes = config.get("supervisor_alert_jefes_group", "")
        if grupo_jefes:
            msg_jefes = (
                f"Hola a todos, siendo las {hora_lima} aun no llega la info de Integratel, "
                f"seguimos a la espera. Si llega se les enviara."
            )
            wa.send_text(grupo_jefes, msg_jefes)
            logging.info(f"[OK] Aviso a jefes enviado a '{grupo_jefes}'")
        jose = config.get("jose_trigger_contact", "Jose Huiza")
        msg_jose = pick_variant(
            variants.get("jose_trigger_alerta", []),
            "Hola Jose, no llego el correo del avance Movistar. Podrias revisar el proceso cuando puedas? Gracias.",
        )
        wa.send_text(jose, msg_jose)
        logging.info(f"[OK] Aviso a Jose enviado a '{jose}'")
        return True
    except Exception as e:
        logging.warning(f"No se pudo enviar avisos: {e}")
        return False


def _avisar_trigger_llego() -> bool:
    try:
        from msg_utils import pick_variant
        wa, config = _wa_client_y_config()
        grupo_jefes = config.get("supervisor_alert_jefes_group", "")
        if not grupo_jefes:
            return True
        variants = config.get("message_variants", {})
        msg = pick_variant(
            variants.get("jefes_trigger_llego", []),
            "Hola a todos, el avance de ventas ya esta disponible. Disculpen la demora, se envia a continuacion.",
        )
        wa.send_text(grupo_jefes, msg)
        logging.info(f"[OK] Aviso 'ya llego' enviado a '{grupo_jefes}'")
        return True
    except Exception as e:
        logging.warning(f"No se pudo avisar 'ya llego': {e}")
        return False


# ── Verificacion real del estado del dia ──────────────────────────────────────

def _periodo_desde_config() -> str:
    """Lee el periodo YYYY-MM del config.json del proyecto."""
    try:
        cfg_path = BASE_DIR / "config.json"
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        return cfg.get("periodo", "")
    except Exception:
        return ""


def _verificar_archivos() -> dict:
    """
    Verifica si los archivos de salida del dia existen y tienen contenido.
    Retorna un dict con claves: avance_ok, resumido_ok, seguimiento_ok, detalles.
    """
    resultado = {"avance_ok": False, "resumido_ok": False, "seguimiento_ok": False, "detalles": []}
    periodo = _periodo_desde_config()  # ej: "2026-05"

    archivos_dir = BASE_DIR / "Archivos_Avance"

    # Archivo AVANCE_{fecha}.xlsx — buscar cualquier AVANCE_YYYY-MM-DD.xlsx del periodo
    if periodo:
        anio, mes = periodo.split("-")
        patron = f"AVANCE_{anio}-{mes}-*.xlsx"
        candidatos = sorted(archivos_dir.glob(patron))
        # Excluir _tmp
        candidatos = [f for f in candidatos if "_tmp" not in f.name]
        if candidatos:
            ultimo = candidatos[-1]
            resultado["avance_ok"] = ultimo.stat().st_size > 10_000
            resultado["detalles"].append(
                f"AVANCE: {ultimo.name} ({ultimo.stat().st_size:,} bytes) — {'OK' if resultado['avance_ok'] else 'VACIO/CORRUPTO'}"
            )
        else:
            resultado["detalles"].append(f"AVANCE: no encontrado para periodo {periodo}")
    else:
        resultado["detalles"].append("AVANCE: no se pudo determinar periodo")

    # Archivo AVANCE_RESUMIDO.xlsx
    resumido = archivos_dir / "AVANCE_RESUMIDO.xlsx"
    if not resumido.exists():
        # Buscar con nombre alternativo
        resumidos = sorted(archivos_dir.glob("AVANCE_RESUMIDO*.xlsx"))
        resumido = resumidos[-1] if resumidos else None

    if resumido and resumido.exists():
        # Verificar que fue modificado hoy
        mtime = datetime.fromtimestamp(resumido.stat().st_mtime)
        es_hoy = mtime.date() == date.today()
        resultado["resumido_ok"] = es_hoy and resumido.stat().st_size > 5_000
        resultado["detalles"].append(
            f"RESUMIDO: {resumido.name} modificado {mtime.strftime('%H:%M')} — {'OK' if resultado['resumido_ok'] else 'DESACTUALIZADO o VACIO'}"
        )
    else:
        resultado["detalles"].append("RESUMIDO: no encontrado")

    # Archivo SEGUIMIENTO_VDD_FIJA del dia
    if periodo:
        anio, mes = periodo.split("-")
        # El seguimiento usa fecha del ultimo dia procesado, buscar el mas reciente del mes
        seguimientos = sorted(archivos_dir.glob(f"SEGUIMIENTO_VDD_FIJA_*-{mes}-{anio}.xlsx"))
        if seguimientos:
            ultimo_seg = seguimientos[-1]
            mtime_seg = datetime.fromtimestamp(ultimo_seg.stat().st_mtime)
            es_hoy_seg = mtime_seg.date() == date.today()
            resultado["seguimiento_ok"] = es_hoy_seg and ultimo_seg.stat().st_size > 5_000
            resultado["detalles"].append(
                f"SEGUIMIENTO: {ultimo_seg.name} modificado {mtime_seg.strftime('%H:%M')} — {'OK' if resultado['seguimiento_ok'] else 'DESACTUALIZADO o VACIO'}"
            )
        else:
            resultado["detalles"].append(f"SEGUIMIENTO: no encontrado para {mes}/{anio}")

    return resultado


def _verificar_modulos(estado: dict) -> dict:
    """
    Analiza el execution_log para determinar que modulos no se ejecutaron o fallaron.
    Retorna: {"completo": bool, "faltantes": list[str], "fallidos": list[str]}
    """
    modulos_esperados = ["backs", "jefes", "jesus", "cristian", "guillermo", "carlos", "italo", "supervisor_alert"]
    modulos_log = estado.get("modulos", {})

    faltantes = [m for m in modulos_esperados if m not in modulos_log]
    fallidos  = [m for m, ok in modulos_log.items() if not ok]

    return {
        "completo": len(faltantes) == 0 and len(fallidos) == 0,
        "faltantes": faltantes,
        "fallidos": fallidos,
    }


def _diagnostico_completo(estado: dict) -> dict:
    """
    Consolida la verificacion de archivos + modulos + execution_log.
    Retorna un dict con el diagnostico y la accion recomendada.
    """
    archivos   = _verificar_archivos()
    modulos    = _verificar_modulos(estado)
    avance_ok  = estado.get("avance_ok", False)
    all_sent   = estado.get("all_sent", False)

    # Determinar accion
    if all_sent and archivos["avance_ok"] and modulos["completo"]:
        # Log dice enviado, archivos OK, modulos OK — nada que hacer
        accion = "nada"
    elif archivos["avance_ok"] and archivos["seguimiento_ok"] and not modulos["fallidos"]:
        # Los archivos reales del dia existen y estan frescos, y no hay ningun modulo
        # con fallo explicito registrado. Los "faltantes" en el log pueden deberse a
        # una ejecucion manual que no escribio en el log (caso tipico: run_modulo.py
        # ejecutado antes de que tuviera soporte de execution_log). En este caso
        # confiar en los artefactos reales es mas seguro que relanzar todo.
        accion = "nada"
    elif not avance_ok or not archivos["avance_ok"]:
        # AVANCE.py no genero el archivo — hay que correr todo desde cero
        accion = "todos"
    else:
        # AVANCE.py OK pero hay modulos con fallo explicito
        pendientes = list(set(modulos["fallidos"]))
        accion = pendientes if pendientes else "todos"

    return {
        "accion": accion,
        "archivos": archivos,
        "modulos": modulos,
        "all_sent_log": all_sent,
    }


def _relanzar_run_modulo(modulos) -> bool:
    """
    Ejecuta run_modulo.py con el argumento indicado.
    modulos puede ser "todos" o una lista de nombres.
    """
    script = BASE_DIR / "run_modulo.py"
    if modulos == "todos":
        args = [sys.executable, str(script), "todos"]
        desc = "todos"
    else:
        # run_modulo.py acepta un modulo a la vez; correrlos en secuencia
        resultados = {}
        for m in modulos:
            logging.info(f"  Relanzando modulo: {m}")
            r = subprocess.run(
                [sys.executable, str(script), m],
                capture_output=True, text=True,
                encoding="utf-8", errors="replace",
            )
            for line in (r.stdout or "").strip().splitlines()[-20:]:
                logging.info(f"    [{m}] {line}")
            if r.returncode != 0:
                for line in (r.stderr or "").strip().splitlines()[-10:]:
                    logging.error(f"    [{m}] {line}")
            resultados[m] = (r.returncode == 0)
        ok_count = sum(1 for v in resultados.values() if v)
        logging.info(f"Modulos relanzados: {ok_count}/{len(resultados)} exitosos")
        return all(resultados.values())

    logging.info(f"Relanzando: {' '.join(args)}")
    result = subprocess.run(args, capture_output=True, text=True, encoding="utf-8", errors="replace")
    for line in (result.stdout or "").strip().splitlines():
        logging.info(f"  [run_modulo] {line}")
    if result.returncode != 0:
        for line in (result.stderr or "").strip().splitlines()[-30:]:
            logging.error(f"  [run_modulo] {line}")
    return result.returncode == 0


def _pipeline_activo() -> bool:
    """Devuelve True si main_v2.py o run_modulo.py estan corriendo ahora mismo."""
    try:
        result = subprocess.run(
            ["wmic", "process", "where", "name='python.exe'", "get", "CommandLine", "/FORMAT:CSV"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        cmd_lines = result.stdout.lower()
        for marker in ("main_v2.py", "run_modulo.py", "avance.py"):
            if marker in cmd_lines:
                return True
        return False
    except Exception:
        return False


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    logging.info("=" * 60)
    logging.info(f"WATCHDOG — {HOY} {datetime.now().strftime('%H:%M:%S')}")
    logging.info("=" * 60)

    estado = _leer_estado_hoy()
    logging.info(f"Estado del dia: {estado}")

    trigger_ok = estado.get("trigger_detected", False)

    # ── Guardia: no interferir si el pipeline esta corriendo ──────────────────
    if _pipeline_activo():
        logging.info("Pipeline activo en este momento. Watchdog espera.")
        return

    # ── Sin trigger ───────────────────────────────────────────────────────────
    if not trigger_ok:
        hora_actual = datetime.now().hour
        if hora_actual >= 13:
            ya_avisado = estado.get("jose_avisado", False)
            if ya_avisado:
                logging.info("Sin trigger — Jose ya fue avisado hoy. Nada mas que hacer.")
            else:
                logging.warning("Sin trigger a las 13:00+. Avisando a jefes y a Jose Huiza...")
                ok = _avisar_sin_trigger()
                _actualizar_log({"jose_avisado": ok, "jose_aviso_hora": datetime.now().isoformat()})
        else:
            logging.info("Sin trigger aun (antes de las 13:00) — OK, puede llegar mas tarde.")
        return

    # ── Con trigger: diagnostico real ─────────────────────────────────────────
    diag = _diagnostico_completo(estado)

    logging.info("--- DIAGNOSTICO ---")
    logging.info(f"  all_sent (log):   {diag['all_sent_log']}")
    logging.info(f"  avance_ok (log):  {estado.get('avance_ok', False)}")
    for det in diag["archivos"]["detalles"]:
        logging.info(f"  [archivo] {det}")
    logging.info(f"  modulos faltantes: {diag['modulos']['faltantes']}")
    logging.info(f"  modulos fallidos:  {diag['modulos']['fallidos']}")
    logging.info(f"  accion recomendada: {diag['accion']}")
    logging.info("-------------------")

    if diag["accion"] == "nada":
        logging.info("[OK] Verificacion real: todo completo. Nada que hacer.")
        return

    # ── Verificar limite de reintentos ────────────────────────────────────────
    reintentos = estado.get("watchdog_reintentos", 0)
    if reintentos >= MAX_REINTENTOS:
        logging.error(
            f"[LIMITE] Se alcanzaron {MAX_REINTENTOS} reintentos automaticos hoy. "
            "No se reintenta mas. Revisa los logs manualmente."
        )
        return

    # ── Ejecutar reintento ────────────────────────────────────────────────────
    accion = diag["accion"]
    n = reintentos + 1
    logging.warning(f"Iniciando reintento #{n} — accion: {accion}")
    _actualizar_log({
        "watchdog_reintentos": n,
        f"watchdog_reintento_{n}_hora": datetime.now().isoformat(),
        f"watchdog_reintento_{n}_accion": accion if accion == "todos" else ",".join(accion),
    })

    ok = _relanzar_run_modulo(accion)

    _actualizar_log({f"watchdog_reintento_{n}_ok": ok})

    if ok:
        logging.info(f"[OK] Reintento #{n} exitoso.")
    else:
        logging.error(f"[ERROR] Reintento #{n} fallo. Proxima ejecucion del watchdog intentara de nuevo (max {MAX_REINTENTOS}).")


if __name__ == "__main__":
    main()
