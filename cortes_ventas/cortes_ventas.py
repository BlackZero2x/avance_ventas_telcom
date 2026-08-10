"""
cortes_ventas.py — Informe de cortes horarios de ventas declaradas por supervisores.

Lee el Google Sheet de respuestas del formulario, calcula tablas por ZONAL y SUPERVISOR,
genera imágenes Excel formateadas y las envía a los grupos de WhatsApp correspondientes.

Uso — informe (se ejecuta en cada horario de corte):
    python cortes_ventas.py --corte 12PM
    python cortes_ventas.py --corte 2PM
    python cortes_ventas.py --corte 4PM
    python cortes_ventas.py --corte 6PM
    python cortes_ventas.py --corte CIERRE

Uso — alerta previa (se ejecuta 10 min antes de cada corte):
    python cortes_ventas.py --corte 12PM --alerta
    python cortes_ventas.py --corte CIERRE --alerta
"""
import argparse
import logging
import os
import sys
import time
import tempfile
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote


# ── Configurar proxy y entorno ────────────────────────────────────────────────

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


def _configurar_proxy():
    no_proxy = "localhost,127.0.0.1,.googleapis.com,.google.com,.gstatic.com"
    os.environ["NO_PROXY"] = no_proxy
    os.environ["no_proxy"] = no_proxy

    proxy_host = os.environ.get("HTTP_PROXY", "")
    user = os.environ.get("PROXY_USER", "")
    password = os.environ.get("PROXY_PASS", "")

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


_cargar_dotenv()
_configurar_proxy()

import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

sys.path.insert(0, str(Path(__file__).parent.parent / "whatsapp_server"))
sys.path.insert(0, str(Path(__file__).parent.parent / "modules"))

from wa_client import WhatsAppClient
from shared.screenshot_safe import ScreenshotManager


# ── Constantes ────────────────────────────────────────────────────────────────

CORTES_VALIDOS = ["12PM", "2PM", "4PM", "6PM", "CIERRE"]
PESOS_CORTE = {"12PM": 0.25, "2PM": 0.50, "4PM": 0.75, "6PM": 1.0, "CIERRE": None}

# Orden de columnas de horas en las tablas
HORAS = ["12PM", "2PM", "4PM", "6PM", "CIERRE"]

SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

BASE_DIR = Path(__file__).parent.parent  # raíz del proyecto
SUBDIR   = Path(__file__).parent         # cortes_ventas/
CONFIG_PATH = BASE_DIR / "whatsapp_server" / "config.json"
TOKEN_PATH = BASE_DIR / "token.json"
CREDS_PATH = BASE_DIR / "credentials.json"
LOGS_DIR = BASE_DIR / "logs"
TEMP_DIR = SUBDIR / "temp"              # Excel e imágenes temporales en cortes_ventas/temp/

# ID del Google Sheet de cortes (hoja Respuestas + hoja CUOTAS)
SHEET_ID = os.environ.get("CORTES_SHEET_ID", "1yrzxBxuwUsB0qNyjnpc5AbkQlojPkU4dS4zm1XftSt8")

# ID del Google Sheet de VENTORY (hoja de ventas registradas)
VENTORY_SHEET_ID = os.environ.get("SHEET_ID_VENTORY", "1EGNnQYG51MROVf2tuxKmEqCkw9lfIXQqJBkJke3AITk")
VENTORY_HOJA_VENTAS = "REG_VTAS_BO"  # nombre de la hoja con datos de ventas registradas (gid=1543273197)

# Nombres de grupos en config.json
GRUPO_VPA = "⚡⚡VPA - Auren"
GRUPO_SUPERVISORES = "Canal Fija 2026 Supervisores y Jefes"
GRUPO_GESTION = "Canal Fija 2026 Gestión AUREN"

# ID de Carlos Parra en WA — recibe mención cuando supervisores no reportan a tiempo
CARLOS_WA_ID = "51968035020@c.us"

# Colores para el Excel de imagen (RGB como tuplas)
COLOR_HEADER = (17, 138, 178)      # #118AB2 azul cielo
COLOR_HEADER_FONT = (255, 255, 255)
COLOR_FILA_PAR = (213, 237, 245)   # azul cielo muy claro
COLOR_TOTAL = (17, 138, 178)       # #118AB2 azul cielo (igual que encabezados)
COLOR_REGION = (44, 62, 80)        # #2C3E50 gris antracita para filas de región
COLOR_VERDE = (131, 235, 176)      # #83EBB0 — #06D6A0 al 50% con blanco
COLOR_AMARILLO = (255, 232, 127)   # #FFE87F — #FFD166 al 50% con blanco
COLOR_ROJO = (247, 163, 183)       # #F7A3B7 — #EF476F al 50% con blanco
COLOR_ROJO_TEXTO = (180, 0, 0)     # rojo oscuro para texto cero
COLOR_NEGRO = (0, 0, 0)
COLOR_BLANCO = (255, 255, 255)

# Agrupación de zonales por región (para tabla VPA)
REGIONES = {
    "REGION LIMA":  ["LIMA"],
    "REGION NORTE": ["CHIMBOTE", "TRUJILLO", "NORTE CHICO"],
    "REGION SUR":   ["AREQUIPA", "ILO", "TACNA"],
}


# ── Logging ───────────────────────────────────────────────────────────────────

LOGS_DIR.mkdir(exist_ok=True)
TEMP_DIR.mkdir(exist_ok=True)

# Log local en cortes_ventas/logs/ (además del log central en logs/)
LOCAL_LOGS_DIR = Path(__file__).parent / "logs"
LOCAL_LOGS_DIR.mkdir(exist_ok=True)

_log_fecha = datetime.now().strftime('%Y%m%d')
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(LOGS_DIR / f"cortes_{_log_fecha}.log", encoding="utf-8"),
        logging.FileHandler(LOCAL_LOGS_DIR / f"cortes_{_log_fecha}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)


# ── Autenticación Google ──────────────────────────────────────────────────────

def _autenticar_sheets():
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
    return build("sheets", "v4", credentials=creds)


# ── Lectura del Google Sheet ──────────────────────────────────────────────────

def _parsear_fecha_corte(serie: "pd.Series") -> "pd.Series":
    """
    Parsea FECHA_CORTE tolerando errores de tipeo en el año (ej. 25/05/0026 → 25/05/2026).
    Si el año parseado es < 100, asume que falta el prefijo '20' y lo corrige.
    """
    dt = pd.to_datetime(serie, dayfirst=True, errors="coerce")
    # Corregir años claramente erróneos (< 100 implica que el usuario escribió '0026' en vez de '2026')
    mask_bad = dt.dt.year < 100
    if mask_bad.any():
        dt = dt.copy()
        dt[mask_bad] = dt[mask_bad].apply(
            lambda d: d.replace(year=d.year + 2000) if pd.notna(d) else d
        )
    return dt.dt.date


def _leer_hoja(service, nombre_hoja):
    resultado = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=nombre_hoja,
    ).execute()
    valores = resultado.get("values", [])
    if not valores or len(valores) < 2:
        return pd.DataFrame()
    encabezados = valores[0]
    filas = valores[1:]
    # Normalizar filas con distinta longitud
    filas_norm = [fila + [""] * (len(encabezados) - len(fila)) for fila in filas]
    return pd.DataFrame(filas_norm, columns=encabezados)


def _normalizar_nombre_sup(nombre: str) -> str:
    """Normaliza nombres de supervisores para matching: mayúsculas, sin espacios extra, sin acentos."""
    import unicodedata
    # Remover acentos
    nombre_sin_acentos = ''.join(
        c for c in unicodedata.normalize('NFD', nombre)
        if unicodedata.category(c) != 'Mn'
    )
    # Eliminar espacios dobles/múltiples y normalizar a un solo espacio
    nombre_normalizado = ' '.join(nombre_sin_acentos.split())
    return nombre_normalizado.strip().upper()


def _leer_ventory_ventas(service, fecha_hoy: date) -> dict:
    """
    Lee la hoja de ventas registradas de VENTORY (gid=1543273197).
    Retorna un dict {supervisor: contador_acumulado_hasta_ahora}

    Filtros:
    - DAY = "HOY" (o fecha_hoy en formato DD/MM/YYYY)
    - Vta_Hoy = "Si"
    - Extrae la hora del campo HORA (formato "hh")
    - Solo cuenta hasta la hora actual del día
    """
    try:
        logging.info(f"Leyendo hoja '{VENTORY_HOJA_VENTAS}' de VENTORY...")
        resultado = service.spreadsheets().values().get(
            spreadsheetId=VENTORY_SHEET_ID,
            range=VENTORY_HOJA_VENTAS,
        ).execute()
        valores = resultado.get("values", [])
        if not valores or len(valores) < 2:
            logging.warning(f"Hoja '{VENTORY_HOJA_VENTAS}' en VENTORY vacia o no encontrada")
            return {}

        encabezados = valores[0]
        filas = valores[1:]
        # Normalizar: rellenar con "" si faltan columnas, truncar si sobran
        max_cols = len(encabezados)
        filas_norm = []
        for fila in filas:
            if len(fila) < max_cols:
                fila_norm = fila + [""] * (max_cols - len(fila))
            else:
                fila_norm = fila[:max_cols]
            filas_norm.append(fila_norm)
        df = pd.DataFrame(filas_norm, columns=encabezados)

        if df.empty:
            return {}

        # Buscar índices de columnas clave
        idx_day = next((i for i, col in enumerate(encabezados) if col.upper() == "DAY"), -1)
        idx_sup = next((i for i, col in enumerate(encabezados) if col.upper() == "SUP"), -1)
        idx_hora = next((i for i, col in enumerate(encabezados) if col.upper() == "HORA"), -1)
        idx_vta_hoy = next((i for i, col in enumerate(encabezados) if col.upper() == "VTA_HOY"), -1)

        if idx_day < 0 or idx_sup < 0 or idx_hora < 0 or idx_vta_hoy < 0:
            logging.warning(f"Columnas requeridas no encontradas en VENTORY: DAY={idx_day}, SUP={idx_sup}, HORA={idx_hora}, VTA_HOY={idx_vta_hoy}")
            logging.info(f"Encabezados disponibles: {encabezados}")
            return {}

        # Normalizar fecha a dos formatos para comparación (01/04/2026 y 1/04/2026)
        fecha_formato1 = fecha_hoy.strftime("%d/%m/%Y")  # 22/07/2026
        fecha_formato2 = f"{fecha_hoy.day}/{fecha_hoy.month}/{fecha_hoy.year}"  # 22/7/2026 (sin ceros a la izquierda)

        contador = {}

        logging.info(f"Procesando {len(df)} filas de VENTORY. Buscando DAY='{fecha_formato1}' o '{fecha_formato2}', Vta_Hoy='Sí'")

        for _, fila in df.iterrows():
            day = str(fila.iloc[idx_day]).strip() if idx_day >= 0 else ""
            sup = str(fila.iloc[idx_sup]).strip() if idx_sup >= 0 else ""
            vta_hoy = str(fila.iloc[idx_vta_hoy]).strip() if idx_vta_hoy >= 0 else ""

            # Filtrar por fecha (DAY puede ser "22/07/2026" o "22/7/2026")
            if day != fecha_formato1 and day != fecha_formato2:
                continue

            # Filtrar por Vta_Hoy = "Sí" (con tilde, como aparece en VENTORY)
            if vta_hoy != "Sí":
                continue

            if not sup or sup == "0":
                continue

            # Contar TODAS las ventas del día (sin filtro de hora)
            # Normalizar nombre para matching con tabla de supervisores
            sup_norm = _normalizar_nombre_sup(sup)
            contador[sup] = contador.get(sup, 0) + 1

        logging.info(f"Ventas registradas en VENTORY cargadas (total del día): {len(contador)} supervisores con ventas")
        if contador:
            logging.info(f"Detalle por supervisor (VENTORY): {contador}")
            logging.info(f"Nombres en VENTORY: {list(contador.keys())}")
        return contador

    except Exception as e:
        logging.error(f"Error leyendo VENTORY ventas: {e}")
        return {}


def cargar_datos(service, corte: str, fecha_hoy: date) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna (df_respuestas_filtrado, df_cuotas).
    df_respuestas contiene solo registros del día y corte dado.
    """
    logging.info(f"Leyendo hoja Respuestas del Sheet...")
    df_resp = _leer_hoja(service, "Respuestas")
    logging.info(f"Leyendo hoja CUOTAS del Sheet...")
    df_cuotas = _leer_hoja(service, "CUOTAS")

    if df_resp.empty:
        logging.warning("La hoja Respuestas esta vacia o no tiene datos.")
        return df_resp, df_cuotas

    # Normalizar tipos
    df_resp["VENTA_REGULAR"] = pd.to_numeric(df_resp.get("VENTA_REGULAR", 0), errors="coerce").fillna(0).astype(int)
    df_resp["VENTA_FLEX"] = pd.to_numeric(df_resp.get("VENTA_FLEX", 0), errors="coerce").fillna(0).astype(int)
    df_resp["VENTA_TOTAL"] = df_resp["VENTA_REGULAR"] + df_resp["VENTA_FLEX"]

    # Normalizar CORTE (puede venir con espacios o diferente capitalización)
    df_resp["CORTE"] = df_resp["CORTE"].str.strip().str.upper()

    # Parsear FECHA_CORTE (tolerante a errores de año como 0026 → 2026)
    df_resp["FECHA_CORTE_DT"] = _parsear_fecha_corte(df_resp["FECHA_CORTE"])

    # Filtrar por fecha y corte
    mask = (df_resp["FECHA_CORTE_DT"] == fecha_hoy) & (df_resp["CORTE"] == corte.upper())
    df_filtrado = df_resp[mask].copy()
    logging.info(f"Registros encontrados para {corte} / {fecha_hoy}: {len(df_filtrado)}")

    return df_filtrado, df_cuotas


# ── Cálculo de tablas ─────────────────────────────────────────────────────────

def _cuotas_por_zonal(df_cuotas: pd.DataFrame) -> dict:
    """Devuelve {zonal: cuota_dia} desde la hoja CUOTAS."""
    cuotas = {}
    if df_cuotas.empty:
        return cuotas
    # Columnas esperadas: ZONAL, SUPERVISOR, CUOTA_DIA
    df_cuotas["CUOTA_DIA"] = pd.to_numeric(df_cuotas.get("CUOTA_DIA", 0), errors="coerce").fillna(0)
    for _, fila in df_cuotas.iterrows():
        zonal = str(fila.get("ZONAL", "")).strip()
        if zonal:
            cuotas[zonal] = cuotas.get(zonal, 0) + fila["CUOTA_DIA"]
    return cuotas


def _cuotas_por_supervisor(df_cuotas: pd.DataFrame) -> dict:
    """Devuelve {sup: cuota_dia} desde la hoja CUOTAS (cuotas directas por supervisor)."""
    cuotas = {}
    if df_cuotas.empty:
        return cuotas
    df_cuotas["CUOTA_DIA"] = pd.to_numeric(df_cuotas.get("CUOTA_DIA", 0), errors="coerce").fillna(0)
    for _, fila in df_cuotas.iterrows():
        sup = str(fila.get("SUPERVISOR", "")).strip()
        if sup:
            cuotas[sup] = fila["CUOTA_DIA"]
    return cuotas


def calcular_tabla_zonal(
    df_todos: pd.DataFrame,
    df_cuotas: pd.DataFrame,
    corte_actual: str,
    fecha_hoy: date,
) -> pd.DataFrame:
    """
    Construye la tabla resumen por ZONAL con columnas:
    ZONAL, 12PM, 2PM, 4PM, 6PM, CIERRE, CUOTA, %ALCANCE, PROYECTADO
    """
    cuotas_zonal = _cuotas_por_zonal(df_cuotas)

    # Necesitamos todos los cortes del día (no solo el actual) para construir las columnas
    df_dia = df_todos.copy() if not df_todos.empty else pd.DataFrame(
        columns=["ZONAL", "CORTE", "VENTA_TOTAL", "FECHA_CORTE_DT", "VENTA_REGULAR", "VENTA_FLEX"]
    )

    # Pivot: zonal x corte → suma de ventas
    if not df_dia.empty:
        pivot = df_dia.groupby(["ZONAL", "CORTE"])["VENTA_TOTAL"].sum().unstack(fill_value=0)
    else:
        pivot = pd.DataFrame()

    # Recoger todas las zonales (del Sheet y de CUOTAS)
    zonales = sorted(set(list(pivot.index if not pivot.empty else [])) | set(cuotas_zonal.keys()))

    filas = []
    for zonal in zonales:
        fila = {"ZONAL": zonal}
        for h in HORAS:
            fila[h] = int(pivot.loc[zonal, h]) if (not pivot.empty and zonal in pivot.index and h in pivot.columns) else 0

        suma_cortes = sum(fila[h] for h in HORAS)
        fila["AVANCE_DIA"] = suma_cortes

        cuota = cuotas_zonal.get(zonal, 0)
        fila["CUOTA"] = int(cuota)
        fila["%ALCANCE"] = (suma_cortes / cuota * 100) if cuota > 0 else 0.0

        peso = PESOS_CORTE.get(corte_actual)
        if peso and peso > 0:
            suma_hasta_corte = sum(fila[h] for h in HORAS if h != "CIERRE")
            fila["PROYECTADO"] = int(round(suma_hasta_corte / peso))
        else:
            fila["PROYECTADO"] = int(suma_cortes)

        filas.append(fila)

    df_tabla = pd.DataFrame(filas, columns=["ZONAL"] + HORAS + ["AVANCE_DIA", "CUOTA", "%ALCANCE", "PROYECTADO"])

    # Fila de totales
    totales = {"ZONAL": "TOTAL"}
    for h in HORAS:
        totales[h] = int(df_tabla[h].sum())
    totales["AVANCE_DIA"] = int(df_tabla["AVANCE_DIA"].sum())
    totales["CUOTA"] = int(df_tabla["CUOTA"].sum())
    totales["%ALCANCE"] = (
        totales["AVANCE_DIA"] / totales["CUOTA"] * 100
        if totales["CUOTA"] > 0 else 0.0
    )
    peso = PESOS_CORTE.get(corte_actual)
    if peso and peso > 0:
        totales["PROYECTADO"] = int(round(
            sum(totales[h] for h in HORAS if h != "CIERRE") / peso
        ))
    else:
        totales["PROYECTADO"] = int(totales["AVANCE_DIA"])

    df_tabla = pd.concat([df_tabla, pd.DataFrame([totales])], ignore_index=True)
    return df_tabla


def _mapear_supervisores(df: pd.DataFrame, mapeo: dict) -> pd.DataFrame:
    """Reemplaza nombres de supervisores según mapeo (ej: GOMEZ PALZA → SUPERVISOR MOQUEGUA)."""
    if df.empty or not mapeo:
        return df
    df = df.copy()
    # Aplicar mapeo a todos los posibles campos de supervisor
    for col in ["SUP", "SUPERVISOR"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: mapeo.get(_normalizar_nombre_sup(x), x) if pd.notna(x) else x)
    return df


def calcular_tabla_supervisor(
    df_todos: pd.DataFrame,
    df_cuotas: pd.DataFrame,
    corte_actual: str,
    fecha_hoy: date,
    service=None,
) -> pd.DataFrame:
    """
    Construye la tabla resumen por SUPERVISOR con columnas:
    ZONAL, SUP, 12PM, 2PM, 4PM, 6PM, CIERRE, CUOTA, %ALCANCE, PROYECTADO, VENTORY

    Si se pasa `service` (cliente de Google Sheets), carga de VENTORY el contador acumulado
    de ventas registradas hasta la hora actual (filtros: DAY=HOY, Vta_Hoy=Si).
    """
    # Mapeo de supervisores antiguos a nuevos nombres
    MAPEO_SUPERVISORES = {
        _normalizar_nombre_sup("GOMEZ PALZA CAROLINA MERCEDES"): "SUPERVISOR MOQUEGUA",
    }

    # Aplicar mapeo a los DataFrames de entrada
    df_todos = _mapear_supervisores(df_todos, MAPEO_SUPERVISORES)
    df_cuotas = _mapear_supervisores(df_cuotas, MAPEO_SUPERVISORES)

    cuotas_sup = _cuotas_por_supervisor(df_cuotas)

    # Mapeo de supervisores antiguos a nuevos nombres
    MAPEO_SUPERVISORES = {
        _normalizar_nombre_sup("GOMEZ PALZA CAROLINA MERCEDES"): "SUPERVISOR MOQUEGUA",
    }

    # Cargar contador de ventas registradas en VENTORY (acumulado hasta ahora)
    ventory_ventas = {}
    if service:
        ventory_ventas_raw = _leer_ventory_ventas(service, fecha_hoy)
        # Aplicar mapeo a nombres de supervisores en VENTORY
        for nombre_orig, count in ventory_ventas_raw.items():
            nombre_mapped = MAPEO_SUPERVISORES.get(_normalizar_nombre_sup(nombre_orig), nombre_orig)
            ventory_ventas[nombre_mapped] = ventory_ventas.get(nombre_mapped, 0) + count

    df_dia = df_todos.copy() if not df_todos.empty else pd.DataFrame(
        columns=["ZONAL", "SUP", "CORTE", "VENTA_TOTAL", "FECHA_CORTE_DT"]
    )

    if not df_dia.empty:
        pivot = df_dia.groupby(["ZONAL", "SUP", "CORTE"])["VENTA_TOTAL"].sum().unstack(fill_value=0)
    else:
        pivot = pd.DataFrame()

    # Recoger todos los supervisores (por key único, normalizando mayúsculas)
    # Crear dict {supervisor_norm: (supervisor_original, lista_de_zonales, zonal_de_cuotas)} para evitar duplicados
    sups_dict = {}  # {supervisor_norm_upper: (supervisor_original, [zonal1, zonal2, ...], zonal_cuotas)}

    if not pivot.empty:
        for zonal, sup in pivot.index:
            sup_norm = _normalizar_nombre_sup(sup)
            if sup_norm not in sups_dict:
                sups_dict[sup_norm] = (sup, [], None)
            zonales_list, zonal_cuotas = sups_dict[sup_norm][1], sups_dict[sup_norm][2]
            if zonal not in zonales_list:
                zonales_list.append(zonal)
            sups_dict[sup_norm] = (sup, zonales_list, zonal_cuotas)

    # Agregar supervisores de CUOTAS (y guardar su zonal)
    if not df_cuotas.empty:
        for _, fila in df_cuotas.iterrows():
            s = str(fila.get("SUPERVISOR", "")).strip()
            z = str(fila.get("ZONAL", "")).strip()
            if s:
                s_norm = _normalizar_nombre_sup(s)
                if s_norm not in sups_dict:
                    sups_dict[s_norm] = (s, [], z)
                else:
                    # Actualizar zonal de CUOTAS si existe
                    sup_original, zonales_list, _ = sups_dict[s_norm]
                    sups_dict[s_norm] = (sup_original, zonales_list, z)

    # Agregar supervisores de VENTORY
    for sup_ventory in ventory_ventas.keys():
        sup_norm = _normalizar_nombre_sup(sup_ventory)
        if sup_norm not in sups_dict:
            sups_dict[sup_norm] = (sup_ventory, [], None)

    # Construir lista de (zonal_principal, sup, sup_norm) y ordenar por ZONAL primero, luego por SUP
    sups_con_zonal = []
    for sup_norm, (sup_original, zonales_list, zonal_cuotas) in sups_dict.items():
        # Prioridad: zonales de cortes horarios, luego zonal de CUOTAS, luego SIN ZONAL
        if zonales_list:
            zonal_principal = zonales_list[0]
        elif zonal_cuotas:
            zonal_principal = zonal_cuotas
        else:
            zonal_principal = "SIN ZONAL"
        sups_con_zonal.append((zonal_principal, sup_original, sup_norm))

    # Ordenar por ZONAL (alfabético), luego por SUP (alfabético)
    sups_con_zonal_sorted = sorted(sups_con_zonal, key=lambda x: (x[0].upper(), x[1].upper()))

    logging.info(f"Supervisores únicos encontrados: {len(sups_con_zonal_sorted)}")
    for i, (zonal, sup, _) in enumerate(sups_con_zonal_sorted):
        logging.info(f"  {i+1}. {zonal} | {sup}")

    filas = []
    for zonal_principal, sup, sup_norm in sups_con_zonal_sorted:
        fila = {"ZONAL": zonal_principal, "SUP": sup}

        # Sumar TODOS los cortes de este supervisor en TODAS las zonales
        for h in HORAS:
            suma_h = 0
            if not pivot.empty:
                for zonal, sup_pivot in pivot.index:
                    if _normalizar_nombre_sup(sup_pivot) == sup_norm and h in pivot.columns:
                        suma_h += int(pivot.loc[(zonal, sup_pivot), h])
            fila[h] = suma_h

        suma_cortes = sum(fila[h] for h in HORAS)
        fila["AVANCE_DIA"] = suma_cortes

        cuota = cuotas_sup.get(sup, 0)
        fila["CUOTA"] = int(cuota)
        fila["%ALCANCE"] = (suma_cortes / cuota * 100) if cuota > 0 else 0.0

        peso = PESOS_CORTE.get(corte_actual)
        if peso and peso > 0:
            suma_hasta_corte = sum(fila[h] for h in HORAS if h != "CIERRE")
            fila["PROYECTADO"] = int(round(suma_hasta_corte / peso))
        else:
            fila["PROYECTADO"] = int(suma_cortes)

        # Agregar contador de ventas registradas en VENTORY
        # Buscar coincidencia exacta primero, luego por coincidencia parcial (primeras palabras)
        contador_ventory = ventory_ventas.get(sup, 0)
        if contador_ventory == 0:
            # Intentar matching aproximado: buscar si algún nombre de VENTORY contiene el nombre del supervisor
            sup_norm = _normalizar_nombre_sup(sup)
            for nombre_ventory, count in ventory_ventas.items():
                nombre_ventory_norm = _normalizar_nombre_sup(nombre_ventory)
                # Match si el nombre del supervisor está contenido en el de VENTORY o viceversa
                if sup_norm in nombre_ventory_norm or nombre_ventory_norm in sup_norm:
                    contador_ventory = count
                    break
        fila["VENTORY"] = contador_ventory

        filas.append(fila)

    cols = ["ZONAL", "SUP"] + HORAS + ["AVANCE_DIA", "CUOTA", "%ALCANCE", "PROYECTADO", "VENTORY"]
    df_tabla = pd.DataFrame(filas, columns=cols)

    # Fila de totales
    totales = {"ZONAL": "TOTAL", "SUP": ""}
    for h in HORAS:
        totales[h] = int(df_tabla[h].sum())
    totales["AVANCE_DIA"] = int(df_tabla["AVANCE_DIA"].sum())
    totales["CUOTA"] = int(df_tabla["CUOTA"].sum())
    totales["%ALCANCE"] = (
        totales["AVANCE_DIA"] / totales["CUOTA"] * 100
        if totales["CUOTA"] > 0 else 0.0
    )
    peso = PESOS_CORTE.get(corte_actual)
    if peso and peso > 0:
        totales["PROYECTADO"] = int(round(
            sum(totales[h] for h in HORAS if h != "CIERRE") / peso
        ))
    else:
        totales["PROYECTADO"] = int(totales["AVANCE_DIA"])
    totales["VENTORY"] = int(df_tabla["VENTORY"].sum())

    df_tabla = pd.concat([df_tabla, pd.DataFrame([totales])], ignore_index=True)
    return df_tabla


# ── Tabla VPA con filas de región ────────────────────────────────────────────

def _insertar_filas_region(df_zonal: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve una copia de df_zonal con filas de subtotal de región intercaladas
    antes de cada grupo de zonales.  La fila TOTAL al final no se toca.
    Las zonales sin región conocida se añaden al final sin encabezado de región.
    Añade columna interna '_tipo': 'region' | 'zonal' | 'total' para el formateador.
    """
    cols_num = HORAS + ["AVANCE_DIA", "CUOTA"]
    col_pct = "%ALCANCE"
    col_proy = "PROYECTADO"

    # Separar total del resto
    mask_total = df_zonal["ZONAL"] == "TOTAL"
    df_datos = df_zonal[~mask_total].copy()
    df_total  = df_zonal[mask_total].copy()

    # Normalizar nombres de zonal para comparación (mayúsculas, sin espacios extra)
    def _norm(s):
        return str(s).strip().upper()

    filas_salida = []

    for nombre_region, zonales_region in REGIONES.items():
        zonales_norm = [_norm(z) for z in zonales_region]
        mask = df_datos["ZONAL"].apply(_norm).isin(zonales_norm)
        grupo = df_datos[mask]
        if grupo.empty:
            continue

        # Fila de subtotal de región
        subtotal = {col: int(grupo[col].sum()) for col in cols_num if col in grupo.columns}
        cuota_reg = subtotal.get("CUOTA", 0)
        avance_reg = subtotal.get("AVANCE_DIA", 0)
        subtotal["ZONAL"] = nombre_region
        subtotal[col_pct] = (avance_reg / cuota_reg * 100) if cuota_reg > 0 else 0.0
        # PROYECTADO de la región = suma de proyectados de sus zonales
        subtotal[col_proy] = int(grupo[col_proy].sum()) if col_proy in grupo.columns else 0
        subtotal["_tipo"] = "region"
        if "SUP" in df_zonal.columns:
            subtotal["SUP"] = ""

        filas_salida.append(subtotal)
        for _, fila in grupo.iterrows():
            d = fila.to_dict()
            d["_tipo"] = "zonal"
            filas_salida.append(d)

    # Zonales que no pertenecen a ninguna región definida
    todas_en_regiones = set()
    for zs in REGIONES.values():
        todas_en_regiones.update([_norm(z) for z in zs])
    sin_region = df_datos[~df_datos["ZONAL"].apply(_norm).isin(todas_en_regiones)]
    for _, fila in sin_region.iterrows():
        d = fila.to_dict()
        d["_tipo"] = "zonal"
        filas_salida.append(d)

    # Fila total
    for _, fila in df_total.iterrows():
        d = fila.to_dict()
        d["_tipo"] = "total"
        filas_salida.append(d)

    cols_out = ["_tipo"] + [c for c in df_zonal.columns]
    df_out = pd.DataFrame(filas_salida, columns=cols_out)
    return df_out


# ── Generación de imagen Excel ────────────────────────────────────────────────

def _rgb(r, g, b):
    """Convierte RGB a valor entero para openpyxl."""
    from openpyxl.styles import PatternFill
    return PatternFill(start_color=f"{r:02X}{g:02X}{b:02X}", end_color=f"{r:02X}{g:02X}{b:02X}", fill_type="solid")


def _color_alcance(valor_pct: float):
    """Retorna RGB según semáforo de %ALCANCE."""
    if valor_pct >= 90:
        return COLOR_VERDE
    elif valor_pct >= 70:
        return COLOR_AMARILLO
    else:
        return COLOR_ROJO


def generar_imagen_tabla(df: pd.DataFrame, titulo: str, ruta_png: str, wb_existente=None, nombre_hoja: str = "Tabla") -> tuple[str, str]:
    """
    Genera una hoja Excel formateada dentro de un workbook existente o crea uno nuevo.

    Retorna (ruta_xlsx, nombre_hoja_creada)

    Si df contiene la columna '_tipo' (valores: 'region'|'zonal'|'total'),
    aplica formato especial para las filas de región.
    La captura PNG se realiza en generar_imagenes_tablas() con una sola instancia de Excel.
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    import xlwings as xw
    import subprocess

    ruta_xlsx = ruta_png.replace(".png", ".xlsx")

    # Usar workbook existente o crear uno nuevo
    if wb_existente is not None:
        wb = wb_existente
        # Crear una nueva hoja con el nombre especificado
        if nombre_hoja in wb.sheetnames:
            # Si ya existe, eliminarla
            del wb[nombre_hoja]
        ws = wb.create_sheet(nombre_hoja)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = nombre_hoja

    # Detectar si el df tiene metadatos de tipo de fila
    tiene_tipos = "_tipo" in df.columns
    tipos = df["_tipo"].tolist() if tiene_tipos else None
    # Trabajar con df sin la columna interna
    df_render = df.drop(columns=["_tipo"]) if tiene_tipos else df

    n_cols = len(df_render.columns)

    # ── Título ────────────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    celda_titulo = ws["A1"]
    celda_titulo.value = titulo
    celda_titulo.font = Font(name="Aptos Narrow", bold=True, size=13, color="FFFFFF")
    celda_titulo.fill = _rgb(*COLOR_HEADER)
    celda_titulo.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # ── Encabezados ───────────────────────────────────────────────────────────
    borde_fino = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    # Color amarillo 50% transparencia (amarillo claro con texto negro para VENTORY)
    COLOR_VENTORY_HEADER = (255, 255, 153)  # #FFFF99 amarillo 50% con blanco

    for col_idx, col_nombre in enumerate(df_render.columns, start=1):
        celda = ws.cell(row=2, column=col_idx, value=col_nombre)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_fino

        # Formato especial para encabezado VENTORY
        if col_nombre == "VENTORY":
            celda.font = Font(name="Aptos Narrow", bold=True, size=10, color="000000")
            celda.fill = _rgb(*COLOR_VENTORY_HEADER)
        else:
            celda.font = Font(name="Aptos Narrow", bold=True, size=10, color="FFFFFF")
            celda.fill = _rgb(*COLOR_HEADER)

    ws.row_dimensions[2].height = 18

    # ── Filas de datos ────────────────────────────────────────────────────────
    n_filas_datos = len(df_render)

    for fila_idx, (row_pos, row) in enumerate(df_render.iterrows(), start=3):
        tipo = tipos[fila_idx - 3] if tiene_tipos else ("total" if fila_idx - 3 == n_filas_datos - 1 else "zonal")
        es_total  = (tipo == "total")
        es_region = (tipo == "region")

        if es_total:
            fill_base = _rgb(*COLOR_TOTAL)
        elif es_region:
            fill_base = _rgb(*COLOR_REGION)
        else:
            fill_base = _rgb(*COLOR_FILA_PAR) if (fila_idx % 2 == 0) else None

        for col_idx, col_nombre in enumerate(df_render.columns, start=1):
            valor = row[col_nombre]
            celda = ws.cell(row=fila_idx, column=col_idx)

            if col_nombre == "%ALCANCE":
                celda.value = valor / 100 if isinstance(valor, (int, float)) else 0
                celda.number_format = "0.0%"
                r, g, b = _color_alcance(float(valor))
                celda.fill = _rgb(r, g, b)
                celda.font = Font(name="Aptos Narrow", bold=True, size=9, color="000000")
            else:
                celda.value = int(valor) if isinstance(valor, float) and valor == int(valor) else valor
                if es_total or es_region:
                    color_txt = "FFFFFF"
                else:
                    COLS_CORTE = {"12PM", "2PM", "4PM", "6PM", "CIERRE"}
                    es_cero = col_nombre in COLS_CORTE and celda.value == 0
                    color_txt = f"{COLOR_ROJO_TEXTO[0]:02X}{COLOR_ROJO_TEXTO[1]:02X}{COLOR_ROJO_TEXTO[2]:02X}" if es_cero else "000000"
                celda.font = Font(name="Aptos Narrow", bold=(es_total or es_region), size=9, color=color_txt)
                if fill_base and col_nombre != "%ALCANCE":
                    celda.fill = fill_base

            celda.alignment = Alignment(
                horizontal="center" if col_nombre not in ("ZONAL", "SUP") else "left",
                vertical="center"
            )
            celda.border = borde_fino

        ws.row_dimensions[fila_idx].height = 16 if es_region else 15

    # ── Anchos de columna ─────────────────────────────────────────────────────
    anchos_fijos = {
        "12PM": 7, "2PM": 7, "4PM": 7, "6PM": 7, "CIERRE": 8,
        "AVANCE_DIA": 11, "CUOTA": 8, "%ALCANCE": 10, "PROYECTADO": 11,
    }
    for col_idx, col_nombre in enumerate(df_render.columns, start=1):
        col_letter = get_column_letter(col_idx)
        if col_nombre in ("ZONAL", "SUP"):
            max_len = max(
                (len(str(v)) for v in df_render[col_nombre] if v is not None),
                default=len(col_nombre)
            )
            max_len = max(max_len, len(col_nombre))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 40)
        else:
            ws.column_dimensions[col_letter].width = anchos_fijos.get(col_nombre, 10)

    # Solo guardar si no hay workbook existente (es decir, si es un wb aislado)
    if wb_existente is None:
        wb.save(ruta_xlsx)
        logging.info(f"Excel temporal guardado: {ruta_xlsx}")
    # La captura real se hace en generar_imagenes_tablas() con una sola instancia de Excel
    return (ruta_xlsx, nombre_hoja)


def _capturar_xlsx_a_png(ws_xw, n_filas: int, n_cols: int, ruta_png: str) -> bool:
    """Captura el rango de datos de una hoja xlwings como PNG via CopyPicture."""
    from openpyxl.utils import get_column_letter
    from PIL import ImageGrab, Image
    import win32gui
    import ctypes

    ultima_col = get_column_letter(n_cols)
    rango = f"A1:{ultima_col}{n_filas}"
    xl_range = ws_xw.range(rango)

    # Traer Excel al frente para que CopyPicture funcione
    try:
        app = ws_xw.book.app
        app.api.Visible = True
        hwnd = [0]
        def _cb(h, _):
            if win32gui.IsWindowVisible(h) and "Microsoft Excel" in win32gui.GetWindowText(h):
                hwnd[0] = h
                return False
            return True
        win32gui.EnumWindows(_cb, None)
        if hwnd[0]:
            ctypes.windll.user32.AllowSetForegroundWindow(ctypes.windll.kernel32.GetCurrentProcessId())
            win32gui.ShowWindow(hwnd[0], 9)
            win32gui.SetForegroundWindow(hwnd[0])
        time.sleep(1)
    except Exception:
        pass

    # CopyPicture con hasta 3 reintentos
    for intento in range(3):
        try:
            # Asegurar que el rango está seleccionado
            xl_range.select()
            time.sleep(0.2)
            xl_range.api.CopyPicture(Appearance=1, Format=2)
            time.sleep(1.5)
            img = ImageGrab.grabclipboard()
            if img:
                # Escalar 2.5x para mayor resolución (HD)
                nuevo_ancho = int(img.width * 2.5)
                nuevo_alto  = int(img.height * 2.5)
                img = img.resize((nuevo_ancho, nuevo_alto), resample=Image.LANCZOS)
                img.save(ruta_png, "PNG")
                logging.info(f"Imagen guardada: {ruta_png} ({nuevo_ancho}x{nuevo_alto})")
                return True
        except Exception as e:
            logging.warning(f"CopyPicture intento {intento+1} fallido: {e}")
            time.sleep(1 + intento)

    logging.error(f"No se pudo capturar imagen para {ruta_png}")
    return False


def generar_imagenes_tablas(trabajos: list[tuple]) -> dict:
    """
    Genera múltiples imágenes PNG en UNA SOLA instancia de Excel.

    trabajos: lista de (df, titulo, ruta_png)
    Retorna {ruta_png: True/False}

    Optimización: crea un único libro Excel con múltiples hojas (una por tabla)
    en lugar de abrir/cerrar 3 libros por separado.
    """
    import openpyxl
    import xlwings as xw
    import subprocess

    resultados = {ruta_png: False for _, _, ruta_png in trabajos}

    # Paso 1: crear UN ÚNICO workbook con todas las hojas
    wb = openpyxl.Workbook()
    # Eliminar la hoja por defecto
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    hojas_info = []  # lista de (ruta_xlsx_unico, nombre_hoja, ruta_png, n_filas, n_cols)

    for idx, (df, titulo, ruta_png) in enumerate(trabajos):
        nombre_hoja = ["VPA", "Zonal", "Supervisor"][idx] if idx < 3 else f"Tabla{idx}"

        # Generar hoja dentro del workbook único
        ruta_xlsx_unico, hoja_creada = generar_imagen_tabla(df, titulo, ruta_png, wb_existente=wb, nombre_hoja=nombre_hoja)

        # n_filas incluye título (1) + encabezado (1) + filas de datos; excluir col '_tipo' si existe
        n_cols_render = len(df.columns) - (1 if "_tipo" in df.columns else 0)
        n_filas_total = len(df) + 2  # título + encabezado + datos
        hojas_info.append((ruta_xlsx_unico, hoja_creada, ruta_png, n_filas_total, n_cols_render))

    # Guardar el workbook UNA SOLA VEZ al final
    ruta_xlsx_unico = hojas_info[0][0] if hojas_info else None
    if ruta_xlsx_unico:
        wb.save(ruta_xlsx_unico)
        logging.info(f"Excel único guardado con {len(hojas_info)} hojas: {ruta_xlsx_unico}")

    # Paso 2: abrir Excel UNA SOLA VEZ y capturar todos (con mutex)
    lock = ScreenshotManager("MOVISTAR_CORTES")
    if not lock.adquirir_lock(timeout=120):
        logging.error("[LOCK] No se pudo adquirir lock de captura — otro proceso usa Excel. Omitiendo capturas.")
        return resultados

    app = None
    try:
        subprocess.run(["taskkill", "/f", "/im", "EXCEL.EXE"], capture_output=True)
        time.sleep(1)
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False

        # Abrir el workbook UNA SOLA VEZ
        if ruta_xlsx_unico and os.path.exists(ruta_xlsx_unico):
            wb_xw = None
            try:
                wb_xw = app.books.open(os.path.normpath(ruta_xlsx_unico))

                # Capturar cada hoja
                for ruta_xlsx, nombre_hoja, ruta_png, n_filas, n_cols in hojas_info:
                    try:
                        ws_xw = wb_xw.sheets[nombre_hoja]
                        ws_xw.activate()
                        time.sleep(0.3)
                        resultados[ruta_png] = _capturar_xlsx_a_png(ws_xw, n_filas, n_cols, ruta_png)
                        logging.info(f"Capturada hoja '{nombre_hoja}' → {ruta_png}")
                    except Exception as e:
                        logging.error(f"Error capturando hoja {nombre_hoja}: {e}")

            except Exception as e:
                logging.error(f"Error abriendo workbook único: {e}")
            finally:
                if wb_xw:
                    try:
                        wb_xw.close()
                    except Exception:
                        pass
                # Eliminar archivo DESPUÉS de cerrar
                try:
                    if ruta_xlsx_unico:
                        os.remove(ruta_xlsx_unico)
                except Exception:
                    pass

    except Exception as e:
        logging.error(f"Error iniciando Excel para capturas: {e}")
    finally:
        if app:
            try:
                app.quit()
            except Exception:
                pass
        lock.liberar_lock()

    return resultados


# ── Texto resumen ─────────────────────────────────────────────────────────────

def _emoji_alcance(pct: float) -> str:
    if pct >= 90:
        return "✅"
    elif pct >= 70:
        return "⚠️"
    return "🔴"


def _avance_dia_total(df_zonal: pd.DataFrame) -> int:
    """Retorna el AVANCE_DIA de la fila TOTAL."""
    if "TOTAL" in df_zonal["ZONAL"].values:
        return int(df_zonal[df_zonal["ZONAL"] == "TOTAL"].iloc[0]["AVANCE_DIA"])
    return int(df_zonal["AVANCE_DIA"].sum())


def construir_caption_vpa(df_zonal: pd.DataFrame, corte: str, fecha: date) -> str:
    """Mensaje formal para el grupo VPA (cliente Movistar). Va como caption de la imagen."""
    fecha_str = fecha.strftime("%d/%m/%Y")
    saludo = "Buenos días" if corte == "CIERRE" else "Buenas tardes"
    avance = _avance_dia_total(df_zonal)

    if corte == "CIERRE":
        return (
            f"{saludo},\n"
            f"Remitimos el cierre de ventas del día *{fecha_str}*.\n"
            f"Acumulado: *{avance}* ventas."
        )
    return (
        f"{saludo},\n"
        f"Remitimos el avance de ventas al corte de las *{corte}* del día *{fecha_str}*.\n"
        f"Acumulado: *{avance}* ventas."
    )


def construir_texto_resumen(df_zonal: pd.DataFrame, corte: str, fecha: date) -> str:
    """Mensaje interno para el grupo de supervisores."""
    fecha_str = fecha.strftime("%d/%m/%Y")
    fila_total = df_zonal[df_zonal["ZONAL"] == "TOTAL"].iloc[0] if "TOTAL" in df_zonal["ZONAL"].values else None
    cuota_total = int(fila_total["CUOTA"]) if fila_total is not None else 0
    alcance = float(fila_total["%ALCANCE"]) if fila_total is not None else 0.0
    proyectado = int(fila_total["PROYECTADO"]) if fila_total is not None else 0
    avance = _avance_dia_total(df_zonal)
    emoji = _emoji_alcance(alcance)

    lineas = [
        f"📊 *CORTE {corte} — {fecha_str}*",
        "━━━━━━━━━━━━━━━━━━━━━",
        f"{emoji} *Avance día:* {avance}",
        f"🎯 *Cuota:* {cuota_total} | *%Alcance:* {alcance:.1f}%",
        f"📈 *Proyectado:* {proyectado}",
        "━━━━━━━━━━━━━━━━━━━━━",
    ]
    return "\n".join(lineas)


# ── Envíos WhatsApp ───────────────────────────────────────────────────────────

def enviar_informe(wa: WhatsAppClient, df_zonal: pd.DataFrame, df_sup: pd.DataFrame, corte: str, fecha: date,
                   df_todos_dia: pd.DataFrame = None, df_cuotas: pd.DataFrame = None):
    fecha_str = fecha.strftime("%d/%m/%Y")

    es_domingo = fecha.weekday() == 6  # domingo=6
    es_sabado  = fecha.weekday() == 5  # sábado=5

    # Domingos: no se envía nada a ningún grupo
    if es_domingo:
        logging.info(f"Fecha {fecha_str} es domingo — omitiendo todos los envios de corte {corte}")
        return

    caption_vpa = construir_caption_vpa(df_zonal, corte, fecha)
    texto_sup   = construir_texto_resumen(df_zonal, corte, fecha)

    # Tabla VPA: versión enriquecida con filas de región intercaladas
    df_zonal_vpa = _insertar_filas_region(df_zonal)

    png_vpa   = str(TEMP_DIR / f"corte_{corte}_zonal_vpa.png")
    png_zonal = str(TEMP_DIR / f"corte_{corte}_zonal.png")
    png_sup   = str(TEMP_DIR / f"corte_{corte}_sup.png")

    titulo_vpa   = f"VENTAS POR ZONAL — CORTE {corte} | {fecha_str}"
    titulo_zonal = f"VENTAS POR ZONAL — CORTE {corte} | {fecha_str}"
    titulo_sup   = f"VENTAS POR SUPERVISOR — CORTE {corte} | {fecha_str}"

    # Generar todas las imágenes en una sola instancia de Excel
    resultados = generar_imagenes_tablas([
        (df_zonal_vpa, titulo_vpa,   png_vpa),
        (df_zonal,     titulo_zonal, png_zonal),
        (df_sup,       titulo_sup,   png_sup),
    ])
    ok_vpa   = resultados.get(png_vpa,   False)
    ok_zonal = resultados.get(png_zonal, False)
    ok_sup   = resultados.get(png_sup,   False)

    # ── Grupo VPA: imagen con regiones + caption formal ───────────────────────
    # CIERRE de sábado (ejecutado el domingo) tampoco va a VPA
    if corte == "CIERRE" and es_sabado:
        logging.info(f"CIERRE de sabado — omitiendo envio a {GRUPO_VPA}")
    else:
        logging.info(f"Enviando tabla ZONAL a: {GRUPO_VPA}")
        if ok_vpa:
            wa.send_image(GRUPO_VPA, png_vpa, caption=caption_vpa)
        else:
            wa.send_text(GRUPO_VPA, caption_vpa)
            logging.warning("No se pudo generar imagen VPA — enviando solo texto.")
        time.sleep(8)

    # ── Grupo Supervisores: texto resumen + imagen zonal + imagen supervisores ─
    logging.info(f"Enviando tabla ZONAL + SUPERVISORES a: {GRUPO_SUPERVISORES}")
    wa.send_text(GRUPO_SUPERVISORES, texto_sup)
    time.sleep(5)
    if ok_zonal:
        wa.send_image(GRUPO_SUPERVISORES, png_zonal, caption="")
    else:
        logging.warning("No se pudo generar imagen zonal — enviando solo texto.")
    time.sleep(8)
    if ok_sup:
        wa.send_image(GRUPO_SUPERVISORES, png_sup, caption="")
    else:
        logging.warning("No se pudo generar imagen supervisores — enviando solo texto.")
    time.sleep(8)

    # ── Notificar a Carlos los supervisores que no reportaron ─────────────────
    if df_todos_dia is not None and df_cuotas is not None:
        time.sleep(5)
        notificar_no_reportaron(wa, df_todos_dia, df_cuotas, corte, fecha)

    # Limpiar temporales
    for f in [png_vpa, png_zonal, png_sup]:
        try:
            if os.path.exists(f):
                os.remove(f)
        except Exception:
            pass


# ── Notificación de no-reportaron (post-corte) ───────────────────────────────

def notificar_no_reportaron(
    wa: WhatsAppClient,
    df_todos_dia: pd.DataFrame,
    df_cuotas: pd.DataFrame,
    corte: str,
    fecha: date,
):
    """
    Envía al grupo de supervisores un mensaje mencionando a Carlos Parra
    con la lista de supervisores que NO enviaron su reporte en este corte.
    Si todos reportaron, no envía nada.
    """
    todos_sups = []
    if not df_cuotas.empty:
        for _, fila in df_cuotas.iterrows():
            sup = str(fila.get("SUPERVISOR", "")).strip()
            if sup:
                todos_sups.append(sup)

    ya_reportaron = _supervisores_que_reportaron(df_todos_dia, corte)
    no_reportaron = [s for s in todos_sups if s not in ya_reportaron]

    fecha_str = fecha.strftime("%d/%m/%Y")
    numero_carlos = CARLOS_WA_ID.replace("@c.us", "")

    if not no_reportaron:
        texto = (
            f"Hola @{numero_carlos}, todos los supervisores reportaron a tiempo "
            f"en el corte *{corte}* del *{fecha_str}*. ✅"
        )
        logging.info("Todos los supervisores reportaron — notificando a Carlos.")
        wa.send_mention(GRUPO_SUPERVISORES, texto, [CARLOS_WA_ID])
        return

    lineas = "\n".join(f"  - {sup}" for sup in sorted(no_reportaron))
    texto = (
        f"Hola @{numero_carlos}, los siguientes supervisores no llegaron a reportar a tiempo "
        f"en el corte *{corte}* del *{fecha_str}*:\n\n{lineas}"
    )

    logging.info(
        f"Notificando a Carlos sobre {len(no_reportaron)} supervisores sin reporte en corte {corte}: "
        + ", ".join(sorted(no_reportaron))
    )
    wa.send_mention(GRUPO_SUPERVISORES, texto, [CARLOS_WA_ID])


# ── Alerta de supervisores pendientes ────────────────────────────────────────

def _telefonos_supervisores(df_cuotas: pd.DataFrame) -> dict:
    """
    Devuelve {nombre_sup: "51XXXXXXXXX@c.us"} leyendo la columna TELEFONO de CUOTAS.
    Acepta números con o sin prefijo 51, con o sin guiones/espacios.
    """
    resultado = {}
    if df_cuotas.empty or "TELEFONO" not in df_cuotas.columns:
        return resultado
    for _, fila in df_cuotas.iterrows():
        sup = str(fila.get("SUPERVISOR", "")).strip()
        tel = str(fila.get("TELEFONO", "")).strip()
        if not sup or not tel:
            continue
        # Limpiar: quitar guiones, espacios, paréntesis
        tel = tel.replace("-", "").replace(" ", "").replace("(", "").replace(")", "")
        # Asegurar prefijo 51
        if tel.startswith("51") and len(tel) == 11:
            wa_id = f"{tel}@c.us"
        elif len(tel) == 9:
            wa_id = f"51{tel}@c.us"
        else:
            logging.warning(f"Telefono invalido para {sup}: {tel} — se omitira de menciones")
            continue
        resultado[sup] = wa_id
    return resultado


def _supervisores_que_reportaron(df_todos_dia: pd.DataFrame, corte: str) -> set:
    """Devuelve el conjunto de nombres de supervisores que ya enviaron el corte dado."""
    if df_todos_dia.empty:
        return set()
    mask = df_todos_dia["CORTE"] == corte.upper()
    return set(df_todos_dia.loc[mask, "SUP"].str.strip().unique())


def alertar_pendientes(
    wa: WhatsAppClient,
    service,
    corte: str,
    fecha: date,
):
    """
    Lee el Sheet, detecta supervisores que NO enviaron el corte indicado
    y envía un mensaje con menciones al grupo de supervisores.
    """
    logging.info(f"Verificando pendientes para alerta previa — corte {corte}")

    df_cuotas = _leer_hoja(service, "CUOTAS")
    df_todos_dia = cargar_todos_del_dia(service, fecha)

    # Todos los supervisores conocidos (de CUOTAS)
    todos_sups = {}
    if not df_cuotas.empty:
        for _, fila in df_cuotas.iterrows():
            sup = str(fila.get("SUPERVISOR", "")).strip()
            if sup:
                todos_sups[sup] = sup

    telefonos = _telefonos_supervisores(df_cuotas)
    ya_reportaron = _supervisores_que_reportaron(df_todos_dia, corte)

    pendientes = [s for s in todos_sups if s not in ya_reportaron]

    logging.info(f"Supervisores totales: {len(todos_sups)} | Ya reportaron: {len(ya_reportaron)} | Pendientes: {len(pendientes)}")

    if not pendientes:
        logging.info("Todos los supervisores ya reportaron. No se envia alerta.")
        return

    # Construir mensaje con menciones
    fecha_str = fecha.strftime("%d/%m/%Y")
    lineas_pendientes = []
    ids_menciones = []
    for sup in sorted(pendientes):
        wa_id = telefonos.get(sup)
        if wa_id:
            # WhatsApp renderiza @número en el mensaje, el nombre se muestra por la mención
            numero = wa_id.replace("@c.us", "")
            lineas_pendientes.append(f"  • @{numero}")
            ids_menciones.append(wa_id)
        else:
            # Sin teléfono: solo texto plano
            lineas_pendientes.append(f"  • {sup}")

    texto = (
        f"⚠️ *ALERTA CORTE {corte} — {fecha_str}*\n"
        f"━━━━━━━━━━━━━━━━━━━━━\n"
        f"Los siguientes supervisores aún *no han enviado* su reporte de corte *{corte}*:\n\n"
        + "\n".join(lineas_pendientes)
        + f"\n\n🕐 El informe se enviará en *10 minutos*.\n"
        f"Por favor reporten su corte a la brevedad. ¡Gracias!"
    )

    logging.info(f"Enviando alerta a grupo supervisores con {len(ids_menciones)} menciones")
    if ids_menciones:
        wa.send_mention(GRUPO_SUPERVISORES, texto, ids_menciones)
    else:
        wa.send_text(GRUPO_SUPERVISORES, texto)


# ── Carga de todos los registros del día ─────────────────────────────────────

def cargar_todos_del_dia(service, fecha_hoy: date) -> pd.DataFrame:
    """
    Carga TODOS los registros del día (todos los cortes) para poder
    construir las columnas de cada hora correctamente.
    """
    logging.info("Leyendo todos los registros del dia...")
    df_resp = _leer_hoja(service, "Respuestas")
    if df_resp.empty:
        return df_resp

    df_resp["VENTA_REGULAR"] = pd.to_numeric(df_resp.get("VENTA_REGULAR", 0), errors="coerce").fillna(0).astype(int)
    df_resp["VENTA_FLEX"] = pd.to_numeric(df_resp.get("VENTA_FLEX", 0), errors="coerce").fillna(0).astype(int)
    df_resp["VENTA_TOTAL"] = df_resp["VENTA_REGULAR"] + df_resp["VENTA_FLEX"]
    df_resp["CORTE"] = df_resp["CORTE"].str.strip().str.upper()
    df_resp["FECHA_CORTE_DT"] = _parsear_fecha_corte(df_resp["FECHA_CORTE"])

    return df_resp[df_resp["FECHA_CORTE_DT"] == fecha_hoy].copy()


# ── Entry point ───────────────────────────────────────────────────────────────

def _es_feriado(fecha: date) -> bool:
    """Verifica si la fecha está en la lista de feriados (variable CORTES_FERIADOS_FECHAS del .env)."""
    feriados_str = os.environ.get("CORTES_FERIADOS_FECHAS", "")
    if not feriados_str:
        return False
    feriados = [datetime.strptime(f.strip(), "%Y-%m-%d").date() for f in feriados_str.split(",") if f.strip()]
    return fecha in feriados


def main():
    parser = argparse.ArgumentParser(description="Informe de cortes de ventas por WhatsApp")
    parser.add_argument(
        "--corte",
        required=True,
        choices=CORTES_VALIDOS,
        help="Corte a informar: 12PM, 2PM, 4PM, 6PM, CIERRE",
    )
    parser.add_argument(
        "--fecha",
        default=None,
        help="Fecha en formato DD/MM/YYYY (por defecto: hoy). Para CIERRE usar la fecha del dia anterior.",
    )
    parser.add_argument(
        "--alerta",
        action="store_true",
        help="Modo alerta: notifica al grupo de supervisores pendientes (se ejecuta 10 min antes del corte).",
    )
    parser.add_argument(
        "--solo-generar",
        action="store_true",
        help="Solo genera las imágenes en temp/ sin enviar nada por WhatsApp. Útil para verificar el resultado visualmente.",
    )
    args = parser.parse_args()
    corte = args.corte.upper()

    from datetime import timedelta
    if args.fecha:
        fecha = datetime.strptime(args.fecha, "%d/%m/%Y").date()
    elif corte == "CIERRE":
        # CIERRE: supervisores reportan al día siguiente → usamos fecha de ayer
        fecha = date.today() - timedelta(days=1)
        logging.info(f"Corte CIERRE: usando fecha de ayer ({fecha})")
    else:
        fecha = date.today()

    # ── Validar si es feriado ────────────────────────────────────────────────────
    if _es_feriado(fecha):
        logging.warning(f"FERIADO DETECTADO ({fecha}) — script deshabilitado. No se ejecutará nada.")
        sys.exit(0)

    # ── Modo solo-generar: no necesita WhatsApp ───────────────────────────────
    if args.solo_generar:
        logging.info("=" * 60)
        logging.info(f"MODO SOLO-GENERAR — {corte} | {fecha.strftime('%d/%m/%Y')}")
        logging.info("=" * 60)
        service = _autenticar_sheets()
        logging.info("Leyendo CUOTAS...")
        df_cuotas = _leer_hoja(service, "CUOTAS")
        df_todos_dia = cargar_todos_del_dia(service, fecha)
        n_corte_actual = len(df_todos_dia[df_todos_dia["CORTE"] == corte]) if not df_todos_dia.empty else 0
        logging.info(f"Registros del dia: {len(df_todos_dia)} | Corte {corte}: {n_corte_actual}")
        df_zonal = calcular_tabla_zonal(df_todos_dia, df_cuotas, corte, fecha)
        df_sup   = calcular_tabla_supervisor(df_todos_dia, df_cuotas, corte, fecha, service=service)
        fecha_str = fecha.strftime("%d/%m/%Y")
        df_zonal_vpa = _insertar_filas_region(df_zonal)
        png_vpa   = str(TEMP_DIR / f"corte_{corte}_zonal_vpa.png")
        png_zonal = str(TEMP_DIR / f"corte_{corte}_zonal.png")
        png_sup   = str(TEMP_DIR / f"corte_{corte}_sup.png")
        generar_imagenes_tablas([
            (df_zonal_vpa, f"VENTAS POR ZONAL — CORTE {corte} | {fecha_str}", png_vpa),
            (df_zonal,     f"VENTAS POR ZONAL — CORTE {corte} | {fecha_str}", png_zonal),
            (df_sup,       f"VENTAS POR SUPERVISOR — CORTE {corte} | {fecha_str}", png_sup),
        ])
        logging.info("Imagenes guardadas en temp/ — no se envio nada por WhatsApp.")
        logging.info(f"  VPA:         {png_vpa}")
        logging.info(f"  Zonal:       {png_zonal}")
        logging.info(f"  Supervisor:  {png_sup}")
        return

    # Conectar WhatsApp (necesario tanto para alerta como para informe)
    import json
    with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
        config = json.load(f)

    wa = WhatsAppClient(
        host=config.get("wa_host", "localhost"),
        port=config.get("wa_port", 8002),
        config_path=str(CONFIG_PATH),
    )
    # Esperar hasta 5 min a que el servidor WA esté listo (cubre reinicio de Chrome que puede tardar varios minutos)
    _WA_MAX_ESPERA = 300
    _wa_espera = 0
    while not wa.is_ready():
        if _wa_espera == 0:
            logging.warning("Servidor WhatsApp no disponible — esperando hasta %ds...", _WA_MAX_ESPERA)
        if _wa_espera >= _WA_MAX_ESPERA:
            logging.error(
                "Servidor WhatsApp no respondio en %ds. "
                "Verifica que wa_server.js este corriendo y que Chrome no este bloqueado.",
                _WA_MAX_ESPERA,
            )
            sys.exit(1)
        time.sleep(15)
        _wa_espera += 15
    logging.info("[OK] WhatsApp listo")

    # Autenticar Sheets
    service = _autenticar_sheets()

    # ── Modo alerta (10 min antes del corte) ─────────────────────────────────
    if args.alerta:
        logging.info("=" * 60)
        logging.info(f"ALERTA PREVIA — {corte} | {fecha.strftime('%d/%m/%Y')}")
        logging.info("=" * 60)
        alertar_pendientes(wa, service, corte, fecha)
        logging.info("ALERTA PREVIA COMPLETADA")
        return

    # ── Modo informe (en el horario del corte) ────────────────────────────────
    logging.info("=" * 60)
    logging.info(f"INFORME CORTES — {corte} | {fecha.strftime('%d/%m/%Y')}")
    logging.info("=" * 60)

    # Leer cuotas
    logging.info("Leyendo CUOTAS...")
    df_cuotas = _leer_hoja(service, "CUOTAS")

    # Leer todos los registros del día
    df_todos_dia = cargar_todos_del_dia(service, fecha)
    n_corte_actual = len(df_todos_dia[df_todos_dia["CORTE"] == corte]) if not df_todos_dia.empty else 0
    logging.info(f"Registros del dia: {len(df_todos_dia)} | Corte {corte}: {n_corte_actual}")

    if n_corte_actual == 0:
        logging.warning(f"No hay registros para el corte {corte} del {fecha}. Continuando con ceros.")

    # Calcular tablas
    df_zonal = calcular_tabla_zonal(df_todos_dia, df_cuotas, corte, fecha)
    df_sup = calcular_tabla_supervisor(df_todos_dia, df_cuotas, corte, fecha, service=service)

    logging.info(f"Tabla zonal: {len(df_zonal)} filas | Tabla supervisores: {len(df_sup)} filas")

    # Enviar
    enviar_informe(wa, df_zonal, df_sup, corte, fecha, df_todos_dia=df_todos_dia, df_cuotas=df_cuotas)

    logging.info("=" * 60)
    logging.info("INFORME CORTES COMPLETADO")
    logging.info("=" * 60)


if __name__ == "__main__":
    main()
