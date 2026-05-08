"""
generar_seguimiento_diario.py

Genera la hoja SEGUIMIENTO con la siguiente estructura por vendedor:
  DATOS VDD (cols A-G de VDD1) | RT_D01 ... RT_Dnn | TOTAL RT
                               | ALT_D01 ... ALT_Dnn | TOTAL ALT
                               | CON_D01 ... CON_Dnn | TOTAL CON

Puede usarse como script independiente o importado desde AVANCE.py:
    from generar_seguimiento_diario import agregar_hoja_seguimiento

Uso standalone:
    python generar_seguimiento_diario.py [YYYY-MM]
"""

import os
import re
import sys
import urllib
from datetime import datetime
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

load_dotenv(Path(__file__).parent / ".env")

# ── Constantes de estilo ──────────────────────────────────────────────────────
COLOR_VDD = "1F3864"
COLOR_RT  = "1E6B3C"
COLOR_ALT = "7E3000"
COLOR_CON = "4A235A"
COLOR_TOT = "000000"
COLS_VDD  = ["ZONAL", "SUPERVISOR", "DNI", "VENDEDOR", "ANTIG", "F_INGRESO", "ESQUEMA"]

_thin       = Side(style="thin", color="CCCCCC")
_border     = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_ALT_FILL   = PatternFill("solid", fgColor="F2F2F2")
_EVEN_FILL  = PatternFill("solid", fgColor="FFFFFF")
_FNT        = "Aptos Narrow"
_FNT_SIZE   = 11
COLOR_ALERT = "FFFF00"


def _make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _make_font(bold=True, color="FFFFFF", size=_FNT_SIZE):
    return Font(name=_FNT, bold=bold, color=color, size=size)


# ── Helpers de datos ──────────────────────────────────────────────────────────

def _crear_engine():
    sql_server   = os.environ.get("SQL_SERVER",   r"AUREN22\AUREN")
    sql_database = os.environ.get("SQL_DATABASE", "eAuren")
    sql_user     = os.environ.get("SQL_USER",     "eauren")
    sql_password = os.environ.get("SQL_PASSWORD", "eauren")
    params = urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={sql_server};"
        f"DATABASE={sql_database};"
        f"UID={sql_user};"
        f"PWD={sql_password};"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={params}")


def _pivot_por_dia(df_src, dni_col, fecha_col, val_col=None, prefix=""):
    tmp = df_src.dropna(subset=[dni_col, fecha_col]).copy()
    tmp["dia"] = pd.to_datetime(tmp[fecha_col]).dt.day
    if val_col and val_col in tmp.columns:
        piv = tmp.groupby([dni_col, "dia"])[val_col].sum().unstack(fill_value=0)
    else:
        piv = tmp.groupby([dni_col, "dia"]).size().unstack(fill_value=0)
    piv.columns = [f"{prefix}_D{int(c):02d}" for c in piv.columns]
    piv = piv.reset_index().rename(columns={dni_col: "DNI"})
    piv["DNI"] = piv["DNI"].astype("Int64")
    day_cols = [c for c in piv.columns if c != "DNI"]
    piv[f"TOTAL_{prefix}"] = piv[day_cols].sum(axis=1)
    return piv


def _sort_block(cols, prefix):
    day_cols  = sorted([c for c in cols if c != f"TOTAL_{prefix}"], key=lambda x: int(x.split("_D")[1]))
    total_col = [f"TOTAL_{prefix}"] if f"TOTAL_{prefix}" in cols else []
    return day_cols + total_col


def _construir_resultado(avance_path, periodo, engine):
    """Lee el AVANCE, consulta SQL y devuelve el DataFrame resultado + orden de columnas."""
    # VDD1 — primeras 7 columnas
    vdd1_raw = pd.read_excel(avance_path, sheet_name="VDD1", header=0)
    vdd1 = vdd1_raw.iloc[:, :7].copy()
    vdd1.columns = COLS_VDD
    vdd1["DNI"] = pd.to_numeric(vdd1["DNI"], errors="coerce").astype("Int64")

    # RT
    rt_raw = pd.read_excel(avance_path, sheet_name="RT", header=0)
    rt_raw["DNI_VENDEDOR"] = pd.to_numeric(rt_raw["DNI_VENDEDOR"], errors="coerce").astype("Int64")
    rt_raw["fecha"] = pd.to_datetime(rt_raw["Fecha_de_alta"], errors="coerce").dt.date

    # ALTAS
    altas_raw = pd.read_excel(avance_path, sheet_name="ALTAS", header=0)
    altas_raw["DNI_VENDEDOR"] = pd.to_numeric(altas_raw["DNI_VENDEDOR"], errors="coerce").astype("Int64")
    altas_raw["fecha"] = pd.to_datetime(altas_raw["Fecha_de_alta"], errors="coerce").dt.date

    # SQL consultas
    sql_con = f"""
SELECT
    [fecha_registro],
    [documento_v],
    [zonal_consulta] AS ZONAL,
    [consulta_unica],
    SUM(1) AS CONSULTAS
FROM [eAuren].[dbo].[fija_base_dito_consultas_hoy]
WHERE periodo='{periodo}' AND tipo='INTENCIONES'
GROUP BY
    [fecha_registro],
    [documento_v],
    [zonal_consulta],
    [consulta_unica]
"""
    df_con = pd.read_sql(sql_con, engine)
    df_con["DNI_VENDEDOR"] = pd.to_numeric(df_con["documento_v"], errors="coerce").astype("Int64")
    df_con["fecha"] = pd.to_datetime(df_con["fecha_registro"], errors="coerce").dt.date

    # Pivots
    piv_rt  = _pivot_por_dia(rt_raw,    "DNI_VENDEDOR", "fecha", prefix="RT")
    piv_alt = _pivot_por_dia(altas_raw, "DNI_VENDEDOR", "fecha", prefix="ALT")
    piv_con = _pivot_por_dia(df_con,    "DNI_VENDEDOR", "fecha", val_col="CONSULTAS", prefix="CON")

    # Merge
    resultado = vdd1.copy()
    resultado = resultado.merge(piv_rt,  on="DNI", how="left")
    resultado = resultado.merge(piv_alt, on="DNI", how="left")
    resultado = resultado.merge(piv_con, on="DNI", how="left")

    # Rellenar NaN
    for c in [col for col in resultado.columns if col not in COLS_VDD]:
        if resultado[c].dtype in ("float64", "object"):
            try:
                resultado[c] = resultado[c].fillna(0).astype(int)
            except Exception:
                resultado[c] = resultado[c].fillna(0)

    # Ordenar columnas
    rt_cols  = _sort_block([c for c in resultado.columns if c.startswith("RT_")],  "RT")
    alt_cols = _sort_block([c for c in resultado.columns if c.startswith("ALT_")], "ALT")
    con_cols = _sort_block([c for c in resultado.columns if c.startswith("CON_")], "CON")
    col_order = COLS_VDD + rt_cols + alt_cols + con_cols
    return resultado[col_order], col_order


def _merge_bloque_headers(ws, row, headers):
    start = 1
    prev = headers[0]
    for ci, h in enumerate(headers[1:], start=2):
        if h != prev:
            if ci - 1 > start:
                ws.merge_cells(start_row=row, start_column=start,
                               end_row=row, end_column=ci - 1)
            start = ci
            prev = h
    ws.merge_cells(start_row=row, start_column=start,
                   end_row=row, end_column=len(headers))


# ── Función principal exportable ──────────────────────────────────────────────

def agregar_hoja_seguimiento(wb, avance_path, periodo, engine):
    """
    Agrega la hoja 'VDD2' al workbook openpyxl recibido.
    Llamar desde AVANCE.py pasando el wb del archivo principal o SEGUIMIENTO_VDD_FIJA_.

    Args:
        wb:           openpyxl.Workbook abierto (se le agrega la hoja)
        avance_path:  Path al AVANCE_{fecha}.xlsx ya generado
        periodo:      str YYYY-MM
        engine:       SQLAlchemy engine ya configurado
    """
    print("[OK] Generando hoja VDD2...")
    resultado, col_order = _construir_resultado(avance_path, periodo, engine)
    print(f"[OK] VDD2: {len(resultado)} vendedores, {len(col_order)} columnas")

    ws = wb.create_sheet("VDD2")

    # ── Índice de columna F_INGRESO (base-1) ────────────────────────
    f_ingreso_ci = COLS_VDD.index("F_INGRESO") + 1  # columna 6

    # ── Columnas extra al final ──────────────────────────────────────
    # Detectar los 3 últimos días disponibles en cada bloque
    rt_day_cols  = sorted([c for c in col_order if c.startswith("RT_D")],
                          key=lambda x: int(x.split("_D")[1]))
    alt_day_cols = sorted([c for c in col_order if c.startswith("ALT_D")],
                          key=lambda x: int(x.split("_D")[1]))
    con_day_cols = sorted([c for c in col_order if c.startswith("CON_D")],
                          key=lambda x: int(x.split("_D")[1]))

    EXTRA_COLS = ["RT_ULT_3D", "CON_ULT_3D", "ALT_ULT_3D", "ALERTAS"]
    col_order_full = col_order + EXTRA_COLS
    n_base = len(col_order)

    # ── Cabeceras fila 1 (columnas) ──────────────────────────────────
    headers_row1, colors_row1 = [], []

    for col in COLS_VDD:
        headers_row1.append(col)
        colors_row1.append(COLOR_VDD)

    for col in [c for c in col_order if c.startswith("RT_")]:
        headers_row1.append(col)
        colors_row1.append(COLOR_TOT if col.startswith("TOTAL") else COLOR_RT)

    for col in [c for c in col_order if c.startswith("ALT_")]:
        headers_row1.append(col)
        colors_row1.append(COLOR_TOT if col.startswith("TOTAL") else COLOR_ALT)

    for col in [c for c in col_order if c.startswith("CON_")]:
        headers_row1.append(col)
        colors_row1.append(COLOR_TOT if col.startswith("TOTAL") else COLOR_CON)

    # Extras: columnas resumen + ALERTAS
    for col in ["RT_ULT_3D", "CON_ULT_3D", "ALT_ULT_3D"]:
        headers_row1.append(col)
        colors_row1.append(COLOR_TOT)

    headers_row1.append("ALERTAS")
    colors_row1.append(COLOR_ALERT)

    # ── Escribir fila 1 (encabezados de columna) ─────────────────────
    for ci, (text, color) in enumerate(zip(headers_row1, colors_row1), start=1):
        cell = ws.cell(row=1, column=ci, value=text)
        cell.fill = _make_fill(color)
        cell.font = _make_font(color="000000" if color == COLOR_ALERT else "FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _border

    # ── Índices de columna de las extras (base-1) ────────────────────
    ci_rt_ult  = n_base + 1
    ci_con_ult = n_base + 2
    ci_alt_ult = n_base + 3
    ci_alertas = n_base + 4

    # Letras de columna de los 3 últimos días de cada bloque
    def _refs_ult3(day_cols, row):
        cols3 = day_cols[-3:] if len(day_cols) >= 3 else day_cols
        refs  = []
        for col_name in cols3:
            idx = col_order.index(col_name) + 1
            refs.append(f"{get_column_letter(idx)}{row}")
        return refs

    # ── Datos fila a fila ────────────────────────────────────────────
    for ri, row_data in enumerate(resultado.itertuples(index=False), start=2):
        fill = _ALT_FILL if ri % 2 == 0 else _EVEN_FILL
        base_font = Font(name=_FNT, size=_FNT_SIZE)

        # Columnas base
        for ci, val in enumerate(row_data, start=1):
            cell = ws.cell(row=ri, column=ci)
            if not isinstance(val, str) and pd.isna(val):
                cell.value = None
            else:
                try:
                    cell.value = int(val)
                except (TypeError, ValueError):
                    cell.value = val
            cell.fill = fill
            cell.border = _border
            cell.font = base_font
            cell.alignment = Alignment(
                horizontal="left" if ci <= len(COLS_VDD) else "center",
                vertical="center"
            )
            # F_INGRESO: formato fecha corta
            if ci == f_ingreso_ci:
                cell.number_format = "DD/MM/YYYY"

        # RT_ULT_3D
        refs_rt = _refs_ult3(rt_day_cols, ri)
        cell_rt = ws.cell(row=ri, column=ci_rt_ult)
        cell_rt.value = f"=SUM({','.join(refs_rt)})" if refs_rt else 0
        cell_rt.fill = fill
        cell_rt.border = _border
        cell_rt.font = base_font
        cell_rt.alignment = Alignment(horizontal="center", vertical="center")
        cell_rt.number_format = "#,##0"

        # CON_ULT_3D
        refs_con = _refs_ult3(con_day_cols, ri)
        cell_con = ws.cell(row=ri, column=ci_con_ult)
        cell_con.value = f"=SUM({','.join(refs_con)})" if refs_con else 0
        cell_con.fill = fill
        cell_con.border = _border
        cell_con.font = base_font
        cell_con.alignment = Alignment(horizontal="center", vertical="center")
        cell_con.number_format = "#,##0"

        # ALT_ULT_3D
        refs_alt = _refs_ult3(alt_day_cols, ri)
        cell_alt = ws.cell(row=ri, column=ci_alt_ult)
        cell_alt.value = f"=SUM({','.join(refs_alt)})" if refs_alt else 0
        cell_alt.fill = fill
        cell_alt.border = _border
        cell_alt.font = base_font
        cell_alt.alignment = Alignment(horizontal="center", vertical="center")
        cell_alt.number_format = "#,##0"

        # ALERTAS — letras de columna para las 3 extras
        col_rt_l  = get_column_letter(ci_rt_ult)
        col_con_l = get_column_letter(ci_con_ult)
        col_alt_l = get_column_letter(ci_alt_ult)
        # TOTAL_ALT: columna en col_order
        total_alt_col = next((c for c in col_order if c == "TOTAL_ALT"), None)
        if total_alt_col:
            ci_total_alt = col_order.index(total_alt_col) + 1
            col_total_alt_l = get_column_letter(ci_total_alt)
            alerta_formula = (
                f'=IF({col_con_l}{ri}=0,"SIN CON ULT3D",'
                f'IF({col_rt_l}{ri}=0,"SIN RT ULT3D",'
                f'IF({col_alt_l}{ri}=0,"SIN ALTAS ULT3D",'
                f'CONCATENATE({col_total_alt_l}{ri}," ALT"))))'
            )
        else:
            alerta_formula = (
                f'=IF({col_con_l}{ri}=0,"SIN CON ULT3D",'
                f'IF({col_rt_l}{ri}=0,"SIN RT ULT3D",'
                f'IF({col_alt_l}{ri}=0,"SIN ALTAS ULT3D","OK")))'
            )
        cell_alerta = ws.cell(row=ri, column=ci_alertas, value=alerta_formula)
        cell_alerta.fill = _make_fill(COLOR_ALERT)
        cell_alerta.border = _border
        cell_alerta.font = Font(name=_FNT, size=_FNT_SIZE, bold=True, color="000000")
        cell_alerta.alignment = Alignment(horizontal="center", vertical="center")

    # ── Anchos de columna ────────────────────────────────────────────
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12
    for ci in range(len(COLS_VDD) + 1, n_base + 1):
        ws.column_dimensions[get_column_letter(ci)].width = (
            11 if col_order[ci - 1].startswith("TOTAL") else 7
        )
    # Extras
    for ci in [ci_rt_ult, ci_con_ult, ci_alt_ult]:
        ws.column_dimensions[get_column_letter(ci)].width = 12
    ws.column_dimensions[get_column_letter(ci_alertas)].width = 18

    ws.row_dimensions[1].height = 25
    ws.freeze_panes = ws.cell(row=2, column=len(COLS_VDD) + 1)

    return ws


# ── Uso standalone ────────────────────────────────────────────────────────────

def _pedir_periodo():
    _default = datetime.now().strftime("%Y-%m")
    if len(sys.argv) > 1:
        _arg = sys.argv[1].strip()
        if re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", _arg):
            return _arg
        print(f"Formato invalido '{_arg}'. Usa YYYY-MM.")
        sys.exit(1)
    return _default


if __name__ == "__main__":
    periodo = _pedir_periodo()
    print(f"[OK] Periodo: {periodo}")

    engine = _crear_engine()
    base_dir    = Path(__file__).parent
    archivos_dir = base_dir / "Archivos_Avance"

    candidatos = list(archivos_dir.glob(f"AVANCE_{periodo}*.xlsx"))
    if not candidatos:
        anio = periodo.split("-")[0]
        candidatos = list(archivos_dir.glob(f"AVANCE_{anio}*.xlsx"))
    if not candidatos:
        print(f"[ERROR] No se encontro AVANCE_{periodo}*.xlsx en {archivos_dir}")
        sys.exit(1)

    avance_path = max(candidatos, key=lambda p: p.stat().st_mtime)
    print(f"[OK] Leyendo: {avance_path.name}")

    wb_out = openpyxl.Workbook()
    # Eliminar hoja vacía por defecto
    wb_out.remove(wb_out.active)

    agregar_hoja_seguimiento(wb_out, avance_path, periodo, engine)
    engine.dispose()

    out_path = base_dir / f"SEGUIMIENTO_DIARIO_{periodo}.xlsx"
    wb_out.save(out_path)
    print(f"[OK] Archivo generado: {out_path}")
