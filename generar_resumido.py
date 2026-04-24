"""
Genera AVANCE_RESUMIDO.xlsx ejecutando directamente las queries SQL
contra AUREN22\AUREN/eAuren, sin necesidad de abrir Excel ni Power Query.

Uso:
    python generar_resumido.py            # usa mes actual (YYYY-MM)
    python generar_resumido.py 2026-04    # periodo especifico
"""
import re
import sys
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import pandas as pd
from sqlalchemy import create_engine

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# ── Configuracion ──────────────────────────────────────────────────────────────

OUTPUT_PATH = r"C:\AVANCE_MOVISTAR\files\AVANCE_RESUMIDO.xlsx"
SERVER      = r"AUREN22\AUREN"
DATABASE    = "eAuren"

SCORING_CASE = """
    CASE
        WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) IS NULL THEN '06.NULL'
        WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) = '0' THEN '06.NULL'
        WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) LIKE 'Fija, aplica FINANCIADO. (UP)%' THEN '03.FINAN UP OBLIG'
        WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) LIKE 'Fija, aplica FINANCIADO. OBLI%' THEN '02.FINAN OBLIG'
        WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) LIKE 'Fija, aplica UPFRONT%' THEN '04.UPFRONT'
        WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) LIKE 'Fija, aplica%FINANCIADO%' THEN '01.FINANCIADO'
        ELSE '06.NULL'
    END AS Scoring
"""

SELECT_ALTAS = """
    t.peticion, t.sub_producto, t.departamento, t.provincia, t.distrito,
    t.vendedor, t.documento_vendedor, t.fecha_registro, t.fecha_venta,
    a.fecha_alta, t.tecnologia_ba, t.velocidad_ba, t.documento,
    [cms_desc_producto] = CASE WHEN a.fecha_alta IS NOT NULL THEN a.desc_estado_peticion ELSE t.desc_estado_peticion END,
    {scoring},
    t.pspaquete_renta_destino, t.cms_codnod, t.cms_nroplano, t.cms_codsrv,
    [pangea] = t.cms_tipreq, t.destinotecnologiatv, t.tipo_uso_destino,
    [Flag_Registro_Unico] = CASE WHEN u.peticion IS NOT NULL THEN 1 ELSE 0 END,
    [Flag_Alta]            = CASE WHEN a.peticion IS NOT NULL THEN 1 ELSE 0 END,
    Cambio_Estado = CONVERT(VARCHAR, t.fechacambioestado, 103),
    Hora_reg      = CONVERT(VARCHAR, t.fec_registro, 108)
""".format(scoring=SCORING_CASE)

SELECT_MIG = """
    t.peticion, t.sub_producto, t.departamento, t.provincia, t.distrito,
    t.vendedor, t.documento_vendedor, t.fecha_registro, t.fecha_venta,
    a.fecha_alta, t.tecnologia_ba, t.velocidad_ba, t.documento,
    [cms_desc_producto] = CASE WHEN a.fecha_alta IS NOT NULL THEN a.desc_estado_peticion ELSE t.desc_estado_peticion END,
    t.cms_desc_producto,
    t.pspaquete_renta_destino, t.cms_codnod, t.cms_nroplano, t.cms_codsrv,
    [pangea] = t.cms_tipreq, t.destinotecnologiatv, t.tipo_uso_destino,
    [Flag_Registro_Unico] = CASE WHEN u.peticion IS NOT NULL THEN 1 ELSE 0 END,
    [Flag_Alta]            = CASE WHEN a.peticion IS NOT NULL THEN 1 ELSE 0 END,
    Cambio_Estado = CONVERT(VARCHAR, t.fechacambioestado, 103),
    Hora_reg      = CONVERT(VARCHAR, t.fec_registro, 108)
"""

FROM_JOINS = """
FROM fija_registros_totales t
    LEFT JOIN fija_altas a             ON t.peticion = a.peticion
    LEFT JOIN fija_registros_unicos u  ON t.peticion = u.peticion
"""


def _pedir_periodo():
    _default = datetime.now().strftime("%Y-%m")
    if len(sys.argv) > 1:
        _arg = sys.argv[1].strip()
        if re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", _arg):
            return _arg
        print(f"  Formato invalido '{_arg}'. Usa YYYY-MM, por ejemplo 2026-04.")
        sys.exit(1)
    return _default


def _header(periodo):
    return f"""
DECLARE @periodo         AS CHAR(7) = '{periodo}';
DECLARE @periodoAnterior AS CHAR(7) = CONVERT(CHAR(7), DATEADD(MONTH, -1, CONVERT(DATE, '{periodo}' + '-01')), 126);
"""


def build_rt_query(periodo, zonal_filter, categoria="'ALTA'"):
    """Registros totales: altas del periodo + pendientes del mes."""
    zonal_sql = f"t.zonal IN ({zonal_filter})" if "," in zonal_filter else f"t.zonal = {zonal_filter}"
    select = SELECT_ALTAS if categoria == "'ALTA'" else SELECT_MIG
    return f"""
{_header(periodo)}
SELECT {select}
{FROM_JOINS}
WHERE {zonal_sql}
  AND t.categoria_producto = {categoria}
  AND (
    (FORMAT(t.fecha_registro,'yyyy-MM') = @periodo         AND FORMAT(a.fecha_alta,'yyyy-MM') = @periodo)
    OR
    (FORMAT(t.fecha_registro,'yyyy-MM') = @periodoAnterior AND FORMAT(a.fecha_alta,'yyyy-MM') = @periodo)
    OR
    (FORMAT(t.fecha_registro,'yyyy-MM') = @periodo         AND a.fecha_alta IS NULL)
  )
ORDER BY t.fecha_registro ASC
"""


def build_altas_query(periodo, zonal_filter, categoria="'ALTA'"):
    """Solo registros con alta confirmada en el periodo."""
    zonal_sql = f"t.zonal IN ({zonal_filter})" if "," in zonal_filter else f"t.zonal = {zonal_filter}"
    select = SELECT_ALTAS if categoria == "'ALTA'" else SELECT_MIG
    return f"""
{_header(periodo)}
SELECT {select}
{FROM_JOINS}
WHERE {zonal_sql}
  AND t.categoria_producto = {categoria}
  AND a.peticion IS NOT NULL
  AND (
    FORMAT(t.fecha_registro,'yyyy-MM') IN (@periodo, @periodoAnterior)
  )
  AND FORMAT(a.fecha_alta,'yyyy-MM') = @periodo
ORDER BY t.fecha_registro ASC
"""


# ── Mapa de hojas ──────────────────────────────────────────────────────────────
#  (nombre_hoja, funcion_query, zonal_filter, categoria)

CHB  = "'CHIMBOTE','HUARAZ','NORTE CHICO'"
AQP  = "'AREQUIPA'"
TCN  = "'TACNA'"
TRU  = "'TRUJILLO'"
ILO  = "'ILO'"

HOJAS = [
    ("RTCHB",       build_rt_query,    CHB, "'ALTA'"),
    ("ALTASCHB",    build_altas_query, CHB, "'ALTA'"),
    ("ALTASAQP",    build_altas_query, AQP, "'ALTA'"),
    ("RTAQP",       build_rt_query,    AQP, "'ALTA'"),
    ("RTTCN",       build_rt_query,    TCN, "'ALTA'"),
    ("ALTASTCN",    build_altas_query, TCN, "'ALTA'"),
    ("RTTRU",       build_rt_query,    TRU, "'ALTA'"),
    ("ALTASTRU",    build_altas_query, TRU, "'ALTA'"),
    ("RTILO",       build_rt_query,    ILO, "'ALTA'"),
    ("ALTASILO",    build_altas_query, ILO, "'ALTA'"),
    ("RTMIGTRU",    build_rt_query,    TRU, "'MIGRACION'"),
    ("MIGTRU",      build_altas_query, TRU, "'MIGRACION'"),
    ("RTMIGAQP",    build_rt_query,    AQP, "'MIGRACION'"),
    ("ALTASMIGAQP", build_altas_query, AQP, "'MIGRACION'"),
    ("RTMIGCHB",    build_rt_query,    CHB, "'MIGRACION'"),
    ("ALTASMIGCHB", build_altas_query, CHB, "'MIGRACION'"),
]


def _run_query(engine, sheet_name, query):
    try:
        df = pd.read_sql(query, engine)
        return sheet_name, df, None
    except Exception as e:
        return sheet_name, None, e


def main():
    periodo = _pedir_periodo()
    logging.info(f"Periodo: {periodo}")

    import urllib
    params = urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        "UID={_sql_user};PWD={_sql_password};"
    )
    engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", fast_executemany=True)

    logging.info(f"Ejecutando {len(HOJAS)} queries en paralelo...")
    resultados = {}
    errores = []

    with ThreadPoolExecutor(max_workers=6) as pool:
        futuros = {
            pool.submit(_run_query, engine, nombre, fn(periodo, zonal, cat)): nombre
            for nombre, fn, zonal, cat in HOJAS
        }
        for futuro in as_completed(futuros):
            nombre, df, err = futuro.result()
            if err:
                logging.error(f"  [ERROR] {nombre}: {err}")
                errores.append(nombre)
                resultados[nombre] = pd.DataFrame()
            else:
                logging.info(f"  [OK] {nombre}: {len(df)} filas")
                resultados[nombre] = df

    INT_COLS = {"peticion", "documento_vendedor"}

    logging.info(f"Escribiendo {OUTPUT_PATH}...")
    DATE_FMT = "DD/MM/YYYY"
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        for nombre, _, _, _ in HOJAS:
            df = resultados.get(nombre, pd.DataFrame())

            # Convertir columnas enteras antes de escribir (NaN → 0 para evitar float)
            for col in INT_COLS:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

            df.to_excel(writer, sheet_name=nombre, index=False)
            ws = writer.sheets[nombre]

            fecha_cols = [
                col_idx + 1
                for col_idx, col_name in enumerate(df.columns)
                if col_name.lower().startswith("fecha")
            ]
            int_col_idxs = [
                col_idx + 1
                for col_idx, col_name in enumerate(df.columns)
                if col_name.lower() in INT_COLS
            ]

            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    if cell.value is None:
                        continue
                    if cell.column in fecha_cols:
                        cell.number_format = DATE_FMT
                    elif cell.column in int_col_idxs:
                        cell.number_format = "0"

    if errores:
        logging.warning(f"Hojas con error: {errores}")
    else:
        logging.info("[OK] AVANCE_RESUMIDO.xlsx generado exitosamente")

    return len(errores) == 0


if __name__ == "__main__":
    ok = main()
    sys.exit(0 if ok else 1)
