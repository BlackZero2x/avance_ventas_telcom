"""
generar_base_datos.py
Genera BASE_DATOS_{mes_anterior}_vs_{mes_actual}.xlsx con 4 hojas:
  - RT:    registros del mes anterior + actual (Fecha_Registro), campo PERIODO en col A
  - ALTAS: altas del mes anterior + actual (Fecha_Alta), campo PERIODO en col A
  - CON:   consultas INTENCIONES de SQL Server (ambos meses)
  - RUS:   Registros Unicos de Servicio de SQL Server (ambos meses)

Los periodos se calculan automaticamente: mes actual vs mes anterior segun la fecha de hoy.

Uso:
    python generar_base_datos.py
"""
import glob
import os
import re
import urllib.parse
from datetime import date
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

load_dotenv(Path(__file__).parent / ".env")

AVANCE_DIR = r"C:\proyectos\AVANCE_MOVISTAR\Archivos_Avance"
OUTPUT_DIR = r"C:\proyectos\AVANCE_MOVISTAR"
# Sin DIA_CORTE aqui: la base guarda todos los dias disponibles del mes.
# El corte lo aplica generar_drill_down.py al leer, usando el ultimo dia con datos.

def _calcular_periodos():
    """Devuelve (periodo_actual, periodo_anterior) en formato 'YYYY-MM' segun la fecha de hoy."""
    hoy = date.today()
    anio_act, mes_act = hoy.year, hoy.month
    if mes_act == 1:
        anio_ant, mes_ant = anio_act - 1, 12
    else:
        anio_ant, mes_ant = anio_act, mes_act - 1
    return f"{anio_act:04d}-{mes_act:02d}", f"{anio_ant:04d}-{mes_ant:02d}"

# Columnas a conservar de las hojas RT y ALTAS (las 322 originales se reducen)
COLS_BASE = [
    "PERIODO",          # campo nuevo, agregado en col A
    "DNI_VENDEDOR",
    "VENDEDOR",
    "DNI",
    "SUPERVISOR",
    "zonal",
    "zonal2",
    "ESQUEMA",
    "Fecha_Registro",
    "Fecha_Alta",
    "producto",
    "sub_producto",
    "segmento",
    "Canal",
    "RIESG",
    "FE",
    "peticion",
]

# Columnas obligatorias segun hoja (las que deben existir para filtrar por fecha)
FECHA_COL = {"RT": "Fecha_Registro", "ALTAS": "Fecha_Alta"}

SQL_CON = """
SELECT
    periodo,
    CONVERT(date, [fecha_registro]) AS fecha_registro,
    [documento_v],
    zonal_consulta AS ZONAL2,
    CASE WHEN zonal_consulta LIKE 'LIMA%' THEN 'LIMA' ELSE zonal_consulta END AS ZONAL,
    SUM(1) AS CON
FROM [eAuren].[dbo].[fija_base_dito_consultas_hoy]
WHERE periodo IN (:p_act, :p_ant) AND tipo='INTENCIONES'
GROUP BY
    periodo,
    CONVERT(date, [fecha_registro]),
    [documento_v],
    [zonal_consulta]
"""

# RUS = Registros Unicos de Servicio (clientes sin pedido en los ultimos 60 dias).
# Es la metrica principal que mide Movistar. Se usa la misma estructura que RT
# pero desde fija_registros_unicos. Se parametriza igual que sql_rt en AVANCE.py.
SQL_RUS_TEMPLATE = """
DECLARE @periodo  AS CHAR(7)
DECLARE @inicio   AS DATE
DECLARE @fin      AS DATE
SET @periodo = '{periodo}';
SET @inicio  = CONVERT(DATE, @periodo + '-01');
SET @fin     = DATEADD(MONTH, 1, @inicio);

WITH realme AS (
    SELECT
        t.peticion,
        tt.producto,
        tt.sub_producto,
        tt.segmento,
        tt.zonal,
        'MASIVO' AS CANAL1,
        CASE
            WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) LIKE 'Fija, aplica FINANCIADO. (UP)%' THEN 'FLEX'
            WHEN t.[cms_desc_tipo_requerimiento] IS NULL THEN 'REGULAR'
            WHEN t.[cms_desc_tipo_requerimiento] = '0' THEN 'REGULAR'
            ELSE 'REGULAR'
        END AS Scoring,
        t.vendedor AS VDD,
        t.documento_vendedor AS DNI_ORIG,
        CAST(t.fecha_registro AS DATE) AS Fecha_Registro,
        CAST(t.fecha_alta AS DATE) AS Fecha_Alta,
        tt.cms_codsrv,
        tt.tecnologia_ba,
        tt.velocidad_ba,
        t.usuario,
        t.destinopaquete
    FROM fija_registros_unicos t
    LEFT JOIN fija_registros_totales tt ON t.peticion = tt.peticion
    WHERE t.categoria_producto = 'ALTA'
      AND t.fecha_registro >= @inicio
      AND t.fecha_registro <  @fin)

SELECT
    @periodo            AS PERIODO,
    peticion,
    VDD,
    DNI_ORIG            AS DNI_VENDEDOR,
    zonal,
    CASE WHEN zonal LIKE 'LIMA%' THEN 'LIMA' ELSE zonal END AS zonal2,
    Fecha_Registro,
    Fecha_Alta,
    producto,
    sub_producto,
    segmento,
    cms_codsrv          AS FE
FROM realme
ORDER BY Fecha_Registro ASC;
"""


def _fecha_key(p):
    m = re.search(r"AVANCE_(\d{4}-\d{2}-\d{2})\.xlsx$", p)
    return m.group(1) if m else "0000-00-00"


def _buscar_ultimo(patron):
    archivos = glob.glob(os.path.join(AVANCE_DIR, patron))
    archivos = [a for a in archivos if re.search(r"AVANCE_\d{4}-\d{2}-\d{2}\.xlsx$", a)]
    if not archivos:
        raise FileNotFoundError(f"No se encontro archivo: {patron}")
    return max(archivos, key=_fecha_key)


def _get_engine():
    params = urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.environ['SQL_SERVER']};"
        f"DATABASE={os.environ['SQL_DATABASE']};"
        f"UID={os.environ['SQL_USER']};"
        f"PWD={os.environ['SQL_PASSWORD']};"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={params}", fast_executemany=True)


def _periodo_del_path(path):
    """Extrae 'YYYY-MM' del nombre del archivo."""
    m = re.search(r"AVANCE_(\d{4}-\d{2})-\d{2}\.xlsx$", os.path.basename(path))
    return m.group(1) if m else None


def _normalizar_zonal2(df):
    """Genera zonal2 si no existe: LIMA para cualquier sub-zonal LIMA, resto igual."""
    if "zonal2" not in df.columns:
        df["zonal2"] = df["zonal"].apply(
            lambda z: "LIMA" if str(z).strip().upper().startswith("LIMA") else str(z).strip().upper()
        )
    return df


def _cargar_hoja(path, hoja, fecha_col, periodo_str):
    """
    Lee hoja del AVANCE, filtra solo el mes del periodo (sin limite de dia),
    agrega campo PERIODO en posicion 0, devuelve solo COLS_BASE disponibles.
    """
    df = pd.read_excel(path, sheet_name=hoja)
    df = _normalizar_zonal2(df)

    df[fecha_col] = pd.to_datetime(df[fecha_col], errors="coerce")
    anio_num = int(periodo_str.split("-")[0])
    mes_num  = int(periodo_str.split("-")[1])

    # Filtrar por anio Y mes para evitar incluir fechas identicas de otro año
    mask = (df[fecha_col].dt.year == anio_num) & (df[fecha_col].dt.month == mes_num)
    df = df[mask].copy()

    # Agregar PERIODO en primera posicion
    df.insert(0, "PERIODO", periodo_str)

    # Conservar solo columnas relevantes que existan en el df
    cols_keep = [c for c in COLS_BASE if c in df.columns]
    # Siempre incluir fecha_col aunque no este en COLS_BASE
    if fecha_col not in cols_keep:
        cols_keep.append(fecha_col)
    return df[cols_keep]


def main():
    periodo_act, periodo_ant = _calcular_periodos()
    print(f"Periodos detectados: actual={periodo_act}  anterior={periodo_ant}")

    path_act = _buscar_ultimo(f"AVANCE_{periodo_act}-*.xlsx")
    path_ant = _buscar_ultimo(f"AVANCE_{periodo_ant}-*.xlsx")

    print(f"Actual  : {os.path.basename(path_act)}  (periodo {periodo_act})")
    print(f"Anterior: {os.path.basename(path_ant)}  (periodo {periodo_ant})")

    engine = _get_engine()

    # ── RT ────────────────────────────────────────────────────────────────────
    print("  Cargando RT...")
    rt_ant = _cargar_hoja(path_ant, "RT", "Fecha_Registro", periodo_ant)
    rt_act = _cargar_hoja(path_act, "RT", "Fecha_Registro", periodo_act)
    rt_all = pd.concat([rt_ant, rt_act], ignore_index=True)
    print(f"    RT {periodo_ant}: {len(rt_ant):,}  |  {periodo_act}: {len(rt_act):,}  |  total: {len(rt_all):,}")

    # ── ALTAS ─────────────────────────────────────────────────────────────────
    print("  Cargando ALTAS...")
    alt_ant = _cargar_hoja(path_ant, "ALTAS", "Fecha_Alta", periodo_ant)
    alt_act = _cargar_hoja(path_act, "ALTAS", "Fecha_Alta", periodo_act)
    alt_all = pd.concat([alt_ant, alt_act], ignore_index=True)
    print(f"    ALTAS {periodo_ant}: {len(alt_ant):,}  |  {periodo_act}: {len(alt_act):,}  |  total: {len(alt_all):,}")

    # ── CON desde SQL ─────────────────────────────────────────────────────────
    print("  Consultando CON desde SQL Server...")
    with engine.connect() as conn:
        con = pd.read_sql(
            text(SQL_CON), conn,
            params={"p_act": periodo_act, "p_ant": periodo_ant}
        )
    con["fecha_registro"] = pd.to_datetime(con["fecha_registro"], errors="coerce")
    con = con.drop(columns=["_dia"], errors="ignore")
    cols_con = ["periodo"] + [c for c in con.columns if c != "periodo"]
    con = con[cols_con].rename(columns={"periodo": "PERIODO"})
    print(f"    CON total registros: {len(con):,}")
    for p in [periodo_ant, periodo_act]:
        n = len(con[con["PERIODO"] == p])
        print(f"      {p}: {n:,}")

    # ── RUS desde SQL (ambos periodos) ────────────────────────────────────────
    # RUS = Registros Unicos de Servicio: clientes sin pedido en los ultimos 60 dias.
    # Es la metrica principal de Movistar. Un vendedor puede tener RT sin RUS
    # (cliente reingresado dentro de los 60 dias, que Movistar no cuenta).
    print("  Consultando RUS desde SQL Server...")
    rus_frames = []
    for periodo_str in [periodo_ant, periodo_act]:
        sql_rus = SQL_RUS_TEMPLATE.format(periodo=periodo_str)
        df_rus = pd.read_sql(sql_rus, engine)
        df_rus["Fecha_Registro"] = pd.to_datetime(df_rus["Fecha_Registro"], errors="coerce")
        df_rus = df_rus.drop(columns=["_dia"], errors="ignore")
        rus_frames.append(df_rus)
        print(f"    RUS {periodo_str}: {len(df_rus):,}")
    rus_all = pd.concat(rus_frames, ignore_index=True)
    print(f"    RUS total: {len(rus_all):,}")

    # ── Guardar ───────────────────────────────────────────────────────────────
    output_path = os.path.join(OUTPUT_DIR, f"BASE_DATOS_{periodo_ant}_vs_{periodo_act}.xlsx")
    print(f"\n  Guardando en {output_path}...")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        rt_all.to_excel(writer,  sheet_name="RT",    index=False)
        alt_all.to_excel(writer, sheet_name="ALTAS", index=False)
        con.to_excel(writer,     sheet_name="CON",   index=False)
        rus_all.to_excel(writer, sheet_name="RUS",   index=False)

    print(
        f"  Listo. Hojas: RT ({len(rt_all):,}), ALTAS ({len(alt_all):,}), "
        f"CON ({len(con):,}), RUS ({len(rus_all):,})"
    )
    # Devolver tambien los periodos para que generar_drill_down los reutilice
    return output_path, periodo_act, periodo_ant


if __name__ == "__main__":
    main()
