"""
Script de pruebas modulares — envia todo a tu propio numero (my_number en config.json).
Uso:
    python test_modulos.py                  # prueba todos los modulos
    python test_modulos.py backs            # solo BacksProcess
    python test_modulos.py jefes
    python test_modulos.py jesus
    python test_modulos.py cristian
    python test_modulos.py italo
    python test_modulos.py wa              # solo prueba conexion WhatsApp
    python test_modulos.py menciones       # prueba mencion real vs texto plano
"""
import json
import logging
import os
import sys
import glob
import time
from datetime import datetime, timedelta

import win32com.client as win32
import pythoncom
import win32clipboard
import win32gui
from PIL import ImageGrab
import xlwings as xw

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "modules"))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "whatsapp_server"))

from wa_client import WhatsAppClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

CONFIG_PATH = "C:/proyectos/AVANCE_MOVISTAR/config.json"

with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
    config = json.load(f)

MY_NUMBER = config.get("my_number", "")  # se toma del config de open-wa
wa = WhatsAppClient(
    host=config.get("wa_host", "localhost"),
    port=config.get("wa_port", 8002)
)

# El numero propio viene del config de open-wa (whatsapp_server/config.json)
WA_CONFIG_PATH = "C:/proyectos/AVANCE_MOVISTAR/whatsapp_server/config.json"
with open(WA_CONFIG_PATH, "r", encoding="utf-8-sig") as f:
    wa_config = json.load(f)
MY_NUMBER = wa_config.get("my_number", "")

if not MY_NUMBER:
    logging.error("No se encontro 'my_number' en whatsapp_server/config.json")
    sys.exit(1)


def buscar_avance():
    """Busca AVANCE_{hoy-1}.xlsx — no usa fallback para evitar enviar datos viejos."""
    directorio = config["archivos_avance_dir"]
    ayer = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    nombre_esperado = os.path.join(directorio, f"AVANCE_{ayer}.xlsx")

    if os.path.exists(nombre_esperado):
        return nombre_esperado

    logging.error(
        f"No se encontro AVANCE_{ayer}.xlsx en {directorio}. "
        f"Ejecuta AVANCE.py primero para generar el archivo del dia."
    )
    return None


def buscar_seguimiento():
    """Busca SEGUIMIENTO_VDD_FIJA_{hoy-1} — no usa fallback para evitar enviar datos viejos."""
    directorio = config["archivos_avance_dir"]
    ayer = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
    nombre_esperado = os.path.join(directorio, f"SEGUIMIENTO_VDD_FIJA_{ayer}.xlsx")

    if os.path.exists(nombre_esperado):
        return nombre_esperado

    logging.error(
        f"No se encontro SEGUIMIENTO_VDD_FIJA_{ayer}.xlsx en {directorio}. "
        f"Ejecuta AVANCE.py primero para generar el archivo del dia."
    )
    return None


# ── Pruebas individuales ───────────────────────────────────────────────────────

def test_wa():
    logging.info("--- PRUEBA: Conexion WhatsApp ---")
    if wa.is_ready():
        logging.info("[OK] Servidor WhatsApp listo")
        wa.send_text(MY_NUMBER, "[PRUEBA] Conexion con open-wa funcionando correctamente.")
    else:
        logging.error("[ERROR] Servidor WhatsApp no disponible")


def test_backs():
    logging.info("--- PRUEBA: BacksProcess (refresh real + carga a Google Sheets) ---")
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from backs_process import BacksProcess

    token_path = config["google_token_path"]
    creds = Credentials.from_authorized_user_file(token_path, [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
    sheets_service = build("sheets", "v4", credentials=creds)

    config_test = dict(config)
    config_test["backs_wa_group"] = MY_NUMBER
    config_test["backs_message"] = f"[PRUEBA BACKS]\n{config['backs_message']}"

    BacksProcess(config_test, sheets_service, wa).execute()
    logging.info("[OK] Prueba BACKS completada")


def _get_excel_hwnd():
    hwnd_found = [0]
    try:
        def _callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "Microsoft Excel" in title or ".xlsx" in title or ".xlsb" in title:
                    hwnd_found[0] = hwnd
                    return False
            return True
        win32gui.EnumWindows(_callback, None)
    except Exception:
        pass
    return hwnd_found[0]


def _capture_range_as_png(wb, sheet_name, cell_range, output_path, excel_hwnd=0):
    MAX_RETRIES = 3
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            sheet = wb.Sheets(sheet_name)
            rng = sheet.Range(cell_range)
            rng.CopyPicture(Appearance=1, Format=2)
            time.sleep(2)  # dar tiempo a COM para llenar el clipboard

            # Intentar abrir clipboard con hwnd de Excel; si falla, usar 0
            opened = False
            for owner in [excel_hwnd, 0]:
                try:
                    win32clipboard.OpenClipboard(owner)
                    opened = True
                    break
                except Exception:
                    time.sleep(0.5)

            if not opened:
                raise RuntimeError("No se pudo abrir el clipboard")

            try:
                img = ImageGrab.grabclipboard()
            finally:
                win32clipboard.CloseClipboard()

            if img is None:
                raise ValueError("Portapapeles vacio tras CopyPicture")

            img.save(output_path, "PNG")
            logging.info(f"  Captura guardada: {output_path}")
            return True
        except Exception as e:
            logging.warning(f"  Intento {attempt}/{MAX_RETRIES} fallido: {e}")
            if attempt < MAX_RETRIES:
                time.sleep(3)
    return False


def test_jefes():
    logging.info("--- PRUEBA: JefesProcess (capturas reales via xlwings) ---")

    archivo = buscar_avance()
    if not archivo:
        logging.error("  No se encontro AVANCE_*.xlsx — abortando prueba jefes")
        return

    temp_dir = config["temp_dir"]
    os.makedirs(temp_dir, exist_ok=True)

    app = None
    wb = None
    try:
        # Cerrar cualquier instancia residual de Excel antes de abrir
        import subprocess
        subprocess.run(["taskkill", "/f", "/im", "EXCEL.EXE"],
                       capture_output=True)
        time.sleep(2)

        logging.info(f"  Abriendo con xlwings: {os.path.basename(archivo)}")
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        wb = app.books.open(os.path.normpath(archivo), update_links=False, read_only=False)

        # Esperar a que Excel esté completamente libre (max 120s)
        logging.info("  Esperando que Excel termine de calcular...")
        for intento in range(120):
            try:
                state = app.api.CalculationState
                if state == 0:  # xlDone
                    logging.info(f"  Excel listo tras {intento}s")
                    break
            except Exception:
                pass
            time.sleep(1)
        else:
            logging.warning("  Excel no termino en 120s, continuando de todas formas")

        app.api.Calculation = -4135  # xlCalculationManual — bloquear nuevos cálculos
        time.sleep(2)

        excel_hwnd = _get_excel_hwnd()
        logging.info(f"  Excel hwnd: {excel_hwnd}")
        sheet = wb.sheets["TDS"]

        for rango, cap_path, msg_key in [
            (config["jefes_tds_rango1"], os.path.join(temp_dir, "prueba_captura1.png"), "jefes_mensaje_captura1"),
            (config["jefes_tds_rango2"], os.path.join(temp_dir, "prueba_captura2.png"), "jefes_mensaje_captura2"),
        ]:
            logging.info(f"  Capturando TDS!{rango}...")
            sheet.range(rango).api.CopyPicture(Appearance=1, Format=2)
            time.sleep(3)  # dar tiempo a Excel para soltar el clipboard

            # Usar PIL.ImageGrab.grabclipboard() directamente — no necesita OpenClipboard manual
            capturado = False
            for intento in range(8):
                try:
                    img = ImageGrab.grabclipboard()
                    if img:
                        img.save(cap_path, "PNG")
                        logging.info(f"  Guardada: {cap_path}")
                        wa.send_image(MY_NUMBER, cap_path, caption=f"[PRUEBA JEFES]\n{config[msg_key]}")
                        time.sleep(2)
                        capturado = True
                        break
                    else:
                        logging.warning(f"  Intento {intento+1}: clipboard vacio")
                        time.sleep(2)
                except Exception as e:
                    logging.warning(f"  Intento {intento+1} grabclipboard: {e}")
                    time.sleep(2)
            if not capturado:
                logging.error(f"  No se pudo capturar TDS!{rango}")

    except Exception as e:
        logging.error(f"  Error con xlwings: {e}")
        import traceback
        logging.error(traceback.format_exc())
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass
        if app:
            try:
                app.quit()
            except Exception:
                pass

    # Archivo SEGUIMIENTO
    seguimiento = buscar_seguimiento()
    if seguimiento:
        logging.info(f"  Enviando SEGUIMIENTO: {os.path.basename(seguimiento)}")
        wa.send_file(MY_NUMBER, seguimiento,
            caption=f"[PRUEBA JEFES - Seguimiento]\n{config['jefes_mensaje_seguimiento']}")
    else:
        logging.warning("  No se encontro SEGUIMIENTO_VDD_FIJA_*.xlsx")

    logging.info("[OK] Prueba JEFES completada")


def test_jesus():
    logging.info("--- PRUEBA: JesusProcess ---")
    archivo = buscar_avance()
    if archivo:
        logging.info(f"  Enviando: {os.path.basename(archivo)}")
        wa.send_file(MY_NUMBER, archivo, caption=f"[PRUEBA JESUS]\n{config['jesus_message']}")
    else:
        logging.warning("  No se encontro AVANCE_*.xlsx — enviando solo mensaje")
        wa.send_text(MY_NUMBER,
            f"[PRUEBA JESUS]\n{config['jesus_message']}\n(archivo no encontrado)"
        )
    logging.info("[OK] Prueba JESUS completada")


def test_cristian():
    logging.info("--- PRUEBA: CristianProcess ---")
    archivo = buscar_avance()
    if archivo:
        logging.info(f"  Enviando: {os.path.basename(archivo)}")
        wa.send_file(MY_NUMBER, archivo, caption=f"[PRUEBA CRISTIAN]\n{config['cristian_message']}")
    else:
        logging.warning("  No se encontro AVANCE_*.xlsx — enviando solo mensaje")
        wa.send_text(MY_NUMBER,
            f"[PRUEBA CRISTIAN]\n{config['cristian_message']}\n(archivo no encontrado)"
        )
    logging.info("[OK] Prueba CRISTIAN completada")


def test_guillermo():
    logging.info("--- PRUEBA: GuillermnoProcess ---")
    archivo = buscar_avance()
    if archivo:
        logging.info(f"  Enviando: {os.path.basename(archivo)}")
        wa.send_file(MY_NUMBER, archivo, caption=f"[PRUEBA GUILLERMO]\n{config['guillermo_message']}")
    else:
        logging.warning("  No se encontro AVANCE_*.xlsx — enviando solo mensaje")
        wa.send_text(MY_NUMBER,
            f"[PRUEBA GUILLERMO]\n{config['guillermo_message']}\n(archivo no encontrado)"
        )
    logging.info("[OK] Prueba GUILLERMO completada")


def test_italo():
    logging.info("--- PRUEBA: ItaloProcess (carga real a Google Sheets) ---")
    import json
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    token_path = config["google_token_path"]
    creds = Credentials.from_authorized_user_file(token_path, [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
    sheets_service = build("sheets", "v4", credentials=creds)

    from italo_process import ItaloProcess
    proceso = ItaloProcess(config, sheets_service, wa)

    # Redirigir el mensaje WA al numero propio
    config_test = dict(config)
    config_test["italo_wa_contact"] = MY_NUMBER
    config_test["italo_message"] = f"[PRUEBA ITALO]\n{config['italo_message']}"

    proceso_test = ItaloProcess(config_test, sheets_service, wa)
    proceso_test.execute()
    logging.info("[OK] Prueba ITALO completada")


def test_carlos():
    logging.info("--- PRUEBA: CarlosProcess (diagnostico de envio) ---")
    archivo = buscar_avance()

    logging.info("TEST 1: Listar contactos que contienen 'Carlos'")
    contactos = wa.list_contacts("Carlos")
    for c in contactos:
        logging.info(f"  Nombre: {c['name']:<30} ID: {c['id']:<20} Push: {c.get('pushname', 'N/A')}")

    logging.info("TEST 2: Enviar texto a 'Carlos P.' (por nombre en config)")
    r1 = wa.send_text(config.get("carlos_wa_contact", "Carlos P."), config.get("carlos_message", "Prueba"))
    logging.info(f"  Resultado: {r1}")
    time.sleep(3)

    if archivo:
        logging.info(f"TEST 3: Enviar archivo a Carlos: {os.path.basename(archivo)}")
        r2 = wa.send_file(config.get("carlos_wa_contact", "Carlos P."), archivo, caption="")
        logging.info(f"  Resultado: {r2}")
    else:
        logging.warning("  No se encontro AVANCE_*.xlsx — omitiendo envio de archivo")

    logging.info("[OK] Prueba CARLOS completada")


def test_menciones():
    """
    Muestra la diferencia entre texto plano con @ y mencion real de WhatsApp.
    La mencion real requiere el ID del contacto (numero@c.us).
    """
    logging.info("--- PRUEBA: Menciones ---")

    # Texto plano (el @ es solo decorativo, no notifica)
    logging.info("  Enviando texto plano con @...")
    wa.send_text(MY_NUMBER,
        "[PRUEBA] Texto plano con @: esto NO genera mencion real en WhatsApp.\n"
        "El @ es solo un caracter de texto."
    )

    # Mencion real usando send_mention (sendTextWithMentions de open-wa)
    logging.info("  Enviando mencion real...")
    wa.send_mention(
        MY_NUMBER,
        f"[PRUEBA] Mencion real: @{MY_NUMBER.replace('@c.us', '')} "
        f"esto SI genera la notificacion de mencion en WhatsApp.",
        mentions=[MY_NUMBER]
    )

    logging.info(
        "[INFO] Para usar menciones en los mensajes del config.json, "
        "configura 'jefes_menciones' con los IDs de WhatsApp:\n"
        '  "jefes_menciones": ["51912345678@c.us", "51987654321@c.us"]\n'
        "Y el texto del mensaje puede incluir @numero para identificar a quien mencionas."
    )
    logging.info("[OK] Prueba MENCIONES completada")


# ── Runner ─────────────────────────────────────────────────────────────────────

def test_orquestador():
    """Ejecuta todos los procesos en orden, sin esperar trigger de email."""
    logging.info("--- ORQUESTADOR MANUAL (sin trigger) ---")
    import os
    _proxy = os.environ.get("PROXY_URL", "http://192.168.2.1:3128")
    try:
        import urllib.request
        urllib.request.urlopen(_proxy, timeout=3)
        os.environ["HTTPS_PROXY"] = _proxy
        os.environ["HTTP_PROXY"]  = _proxy
        os.environ["NO_PROXY"]    = "localhost,127.0.0.1,googleapis.com,google.com"
    except Exception:
        for _v in ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy"):
            os.environ.pop(_v, None)

    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from backs_process import BacksProcess
    from jefes_process import JefesProcess
    from jesus_process import JesusProcess
    from cristian_process import CristianProcess
    from guillermo_process import GuillermnoProcess
    from italo_process import ItaloProcess

    token_path = config["google_token_path"]
    creds = Credentials.from_authorized_user_file(token_path, [
        "https://www.googleapis.com/auth/gmail.modify",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets",
    ])
    sheets_service = build("sheets", "v4", credentials=creds)

    pasos = [
        ("backs",     lambda: BacksProcess(config, sheets_service, wa).execute()),
        ("jefes",     lambda: JefesProcess(config, wa).execute()),
        ("jesus",     lambda: JesusProcess(config, wa).execute()),
        ("cristian",  lambda: CristianProcess(config, wa).execute()),
        ("guillermo", lambda: GuillermnoProcess(config, wa).execute()),
        ("italo",     lambda: ItaloProcess(config, sheets_service, wa).execute()),
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

    ok = sum(1 for v in results.values() if v)
    logging.info("=" * 50)
    logging.info(f"RESUMEN: {ok}/{len(results)} procesos exitosos")
    for nombre, status in results.items():
        logging.info(f"  {'[OK]   ' if status else '[ERROR]'} {nombre}")
    logging.info("=" * 50)


PRUEBAS = {
    "wa":           test_wa,
    "backs":        test_backs,
    "jefes":        test_jefes,
    "jesus":        test_jesus,
    "cristian":     test_cristian,
    "guillermo":    test_guillermo,
    "carlos":       test_carlos,
    "italo":        test_italo,
    "menciones":    test_menciones,
    "orquestador":  test_orquestador,
}

if __name__ == "__main__":
    if not wa.is_ready():
        logging.error("Servidor WhatsApp no disponible. Asegurate de que wa_server.js este corriendo.")
        sys.exit(1)

    modulo = sys.argv[1].lower() if len(sys.argv) > 1 else "todos"

    if modulo == "todos":
        for nombre, fn in PRUEBAS.items():
            if nombre == "menciones":
                continue  # menciones es opt-in, no corre en "todos"
            fn()
    elif modulo in PRUEBAS:
        PRUEBAS[modulo]()
    else:
        print(f"Modulo '{modulo}' no reconocido. Opciones: {', '.join(PRUEBAS.keys())}, todos")
        sys.exit(1)
