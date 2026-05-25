"""
generar_drill_down.py
Genera DRILL_DOWN_ABR_vs_MAY_DD.xlsx con layout basado en ARCHIVO_EJEMPLO_DRILL_DOWN.xlsx.

Hojas: DATA | RESUMEN | FUNNEL | RT | ALT | RUS | CON
Fuente: BASE_DATOS_ABR_MAY.xlsx
"""

import os
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Configuracion ─────────────────────────────────────────────────────────────
OUTPUT_DIR = r"C:\proyectos\AVANCE_MOVISTAR"
FONT_NAME  = "Aptos Narrow"
FONT_SIZE  = 11

ANIO_ABR, MES_ABR = 2026, 4
ANIO_MAY, MES_MAY = 2026, 5

# DIA_CORTE se calcula automaticamente al cargar datos (ver _detectar_dia_corte)
DIA_CORTE: int = 1  # placeholder; se sobreescribe en main()

# Zonales excluidas de todas las vistas (zonales no relevantes o sin asignar)
ZONAL_EXCLUIR = {"HUANUCO", "HUARAZ", "ICA", "SIN ZONAL"}

ZONAL_FULL  = ["LIMA", "AREQUIPA", "CHIMBOTE", "NORTE CHICO",
               "NORTE ORIENTE", "SUR ORIENTE", "SUR", "TACNA", "TRUJILLO", "ILO"]
ZONAL_ABREV = {
    "LIMA":          "LIM",
    "AREQUIPA":      "AQP",
    "CHIMBOTE":      "CHB",
    "NORTE CHICO":   "NCH",
    "NORTE ORIENTE": "NOR",
    "SUR ORIENTE":   "SOR",
    "SUR":           "SUR",
    "TACNA":         "TCN",
    "TRUJILLO":      "TRU",
    "ILO":           "ILO",
}

C = {
    "hdr_dark":  "1A237E",
    "hdr_med":   "283593",
    "hdr_light": "1976D2",
    "sub_hdr":   "BBDEFB",
    "info_hdr":  "E3F2FD",
    "mes_hdr":   "C5CAE9",
    "sem_hdr":   "DCEDC8",
    "dia_hdr":   "FFF9C4",
    "abr":       "FFCC02",
    "may":       "4CAF50",
    "dlt":       "BDBDBD",
    "pct":       "FF8F00",
    "con_row":   "E3F2FD",
    "rt_row":    "E8F5E9",
    "alt_row":   "F3E5F5",
    "rus_row":   "FFF3E0",
    "con_hdr":   "1F5C99",
    "rt_hdr":    "2E7D32",
    "alt_hdr":   "7B1FA2",
    "rus_hdr":   "E65100",
    "pos":       "1B5E20",
    "neg":       "B71C1C",
    "white":     "FFFFFF",
    # semaforo FUNNEL
    "sem_verde":    "C8E6C9",
    "sem_amarillo": "FFF9C4",
    "sem_rojo":     "FFCDD2",
    "sem_verde_fc": "1B5E20",
    "sem_rojo_fc":  "B71C1C",
}

BASE_DATOS_PATH = os.path.join(OUTPUT_DIR, "BASE_DATOS_ABR_MAY.xlsx")

# KPI names usados en DATA sheet
KPI_CON_RATIO = "CON RATIO"
KPI_VDDS_RT   = "VDDS RT"
KPI_VDDS_ALT  = "VDDS ALT"
KPI_VDDS_CON  = "VDDS CON"

KPIS_ORDER = [KPI_CON_RATIO, KPI_VDDS_RT, KPI_VDDS_ALT, KPI_VDDS_CON]


# ── Helpers ────────────────────────────────────────────────────────────────────
def _border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def cs(ws, row, col, value, bold=False, bg=None, fc="000000",
       align="center", num_fmt=None, sz=None, italic=False, border=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=FONT_NAME, bold=bold, color=fc,
                  size=sz or FONT_SIZE, italic=italic)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=False)
    if border:
        c.border = _border()
    if bg:
        c.fill = PatternFill("solid", start_color=bg)
    if num_fmt:
        c.number_format = num_fmt
    return c


def _safe(v):
    if v is None:
        return None
    try:
        f = float(v)
        return None if pd.isna(f) else f
    except Exception:
        return None


def _abrev(z):
    return ZONAL_ABREV.get(z, z[:3])


# ── Deteccion automatica del dia corte ────────────────────────────────────────
def _detectar_dia_corte():
    """
    Lee el maximo dia con datos en RT mayo. Ese es el ultimo dia con informacion
    disponible y se usa como DIA_CORTE para toda la generacion del archivo.
    """
    df = pd.read_excel(BASE_DATOS_PATH, sheet_name="RT",
                       usecols=["PERIODO", "Fecha_Registro"])
    df = df[df["PERIODO"] == f"{ANIO_MAY:04d}-{MES_MAY:02d}"].copy()
    df["dia"] = pd.to_datetime(df["Fecha_Registro"], errors="coerce").dt.day
    dia = int(df["dia"].max()) if not df.empty else 1
    return dia


# ── Dias equivalentes ─────────────────────────────────────────────────────────
def _calcular_dias_equiv():
    """
    Devuelve lista de (d_abr, d_may) para todos los dias de mayo 1..DIA_CORTE.
    d_abr puede ser None si el dia equivalente cae fuera del mes de abril.
    El offset se calcula alineando el primer dia de la semana comun a ambos meses.
    """
    import calendar
    dias_en_abr = calendar.monthrange(ANIO_ABR, MES_ABR)[1]
    abr = [(d, date(ANIO_ABR, MES_ABR, d).weekday()) for d in range(1, dias_en_abr + 1)]
    may = [(d, date(ANIO_MAY, MES_MAY, d).weekday()) for d in range(1, DIA_CORTE + 1)]
    wks_a = {wk: d for d, wk in abr}
    wks_m = {wk: d for d, wk in may}
    comun = sorted(set(wks_a) & set(wks_m))
    if not comun:
        return [(None, d_m) for d_m, _ in may]
    wk0    = comun[0]
    d_a0   = min(d for d, wk in abr if wk == wk0)
    d_m0   = min(d for d, wk in may if wk == wk0)
    offset = d_a0 - d_m0
    pares  = []
    for d_m, _ in may:
        d_a = d_m + offset
        # None si el equivalente cae fuera del mes de abril
        pares.append((d_a if 1 <= d_a <= dias_en_abr else None, d_m))
    return pares


# Se inicializa vacio; main() lo rellena tras detectar DIA_CORTE
DIAS_EQUIV: list = []


def _lbl_dia(d_a, d_m):
    la = date(ANIO_ABR, MES_ABR, d_a).strftime("%d-Abr") if d_a is not None else "s/equiv"
    lm = date(ANIO_MAY, MES_MAY, d_m).strftime("%d-May")
    return f"{la} | {lm}"


# ── Semanas naturales ─────────────────────────────────────────────────────────
def _sem_num(anio, mes, dia):
    dt  = date(anio, mes, dia)
    lu  = dt - timedelta(days=dt.weekday())
    lu1 = date(anio, mes, 1) - timedelta(days=date(anio, mes, 1).weekday())
    return (lu - lu1).days // 7 + 1


def _sem_label(anio, mes, dia):
    return f"S{_sem_num(anio, mes, dia)}"


# ── Carga de datos ────────────────────────────────────────────────────────────
def _filtrar_zonales(df, col="zonal2"):
    """Elimina filas cuya zonal esta en ZONAL_EXCLUIR o es nula."""
    df[col] = df[col].fillna("SIN ZONAL")
    return df[~df[col].isin(ZONAL_EXCLUIR)].copy()


def _pivot_rt_alt(hoja, periodo, fecha_col):
    df = pd.read_excel(BASE_DATOS_PATH, sheet_name=hoja,
                       usecols=["PERIODO", fecha_col, "zonal2"])
    df = df[df["PERIODO"] == periodo].copy()
    df["dia"] = pd.to_datetime(df[fecha_col], errors="coerce").dt.day
    df = df[df["dia"].between(1, DIA_CORTE)]
    df = _filtrar_zonales(df)
    if df.empty:
        return pd.DataFrame()
    pivot = (df.groupby(["dia", "zonal2"]).size()
               .reset_index(name="n")
               .pivot(index="dia", columns="zonal2", values="n")
               .fillna(0).astype(int))
    pivot.index = pivot.index.astype(int)
    pivot = pivot.reindex(range(1, DIA_CORTE + 1), fill_value=0)
    pivot.columns.name = None
    return pivot


def _pivot_vdd_unicos(hoja, periodo, fecha_col):
    """
    Pivot de VDDs unicos POR DIA (para vista diaria y semanal).
    Un vendedor que aparece en 10 dias cuenta 1 vez por dia — correcto para
    la granularidad dia/semana. Para la vista MES se usa _vdd_unicos_mes().
    """
    df = pd.read_excel(BASE_DATOS_PATH, sheet_name=hoja,
                       usecols=["PERIODO", fecha_col, "zonal2", "DNI_VENDEDOR"])
    df = df[df["PERIODO"] == periodo].copy()
    df["dia"] = pd.to_datetime(df[fecha_col], errors="coerce").dt.day
    df = df[df["dia"].between(1, DIA_CORTE)]
    df = _filtrar_zonales(df)
    if df.empty:
        return pd.DataFrame()
    pivot = (df.groupby(["dia", "zonal2"])["DNI_VENDEDOR"]
               .nunique().reset_index(name="n")
               .pivot(index="dia", columns="zonal2", values="n")
               .fillna(0).astype(int))
    pivot.index = pivot.index.astype(int)
    pivot = pivot.reindex(range(1, DIA_CORTE + 1), fill_value=0)
    pivot.columns.name = None
    return pivot


def _vdd_unicos_mes(hoja, periodo, fecha_col):
    """
    Serie zonal -> VDDs unicos del PERIODO COMPLETO (sin sumar por dia).
    Correcto para la fila MES del RESUMEN.
    """
    df = pd.read_excel(BASE_DATOS_PATH, sheet_name=hoja,
                       usecols=["PERIODO", fecha_col, "zonal2", "DNI_VENDEDOR"])
    df = df[df["PERIODO"] == periodo].copy()
    df["dia"] = pd.to_datetime(df[fecha_col], errors="coerce").dt.day
    df = df[df["dia"].between(1, DIA_CORTE)]
    df = _filtrar_zonales(df)
    if df.empty:
        return pd.Series(dtype=int)
    return df.groupby("zonal2")["DNI_VENDEDOR"].nunique()


def _vdd_unicos_mes_con(periodo):
    """VDDs unicos del mes completo para CON (documento_v)."""
    df = pd.read_excel(BASE_DATOS_PATH, sheet_name="CON")
    df["fecha_registro"] = pd.to_datetime(df["fecha_registro"], errors="coerce")
    df["dia"] = df["fecha_registro"].dt.day
    df = df[df["dia"].between(1, DIA_CORTE)].copy()
    df["ZONAL"] = df["ZONAL"].str.strip().str.upper()
    df = _filtrar_zonales(df, col="ZONAL")
    df = df[df["PERIODO"] == periodo]
    if df.empty:
        return pd.Series(dtype=int)
    return df.groupby("ZONAL")["documento_v"].nunique()


def _pivot_rus(periodo):
    """
    Pivot de RUS por (dia, zonal2): conteo de peticiones unicas por dia.
    RUS = Registros Unicos de Servicio, la metrica principal de Movistar.
    """
    df = pd.read_excel(BASE_DATOS_PATH, sheet_name="RUS",
                       usecols=["PERIODO", "Fecha_Registro", "zonal2"])
    df = df[df["PERIODO"] == periodo].copy()
    df["dia"] = pd.to_datetime(df["Fecha_Registro"], errors="coerce").dt.day
    df = df[df["dia"].between(1, DIA_CORTE)]
    df = _filtrar_zonales(df)
    if df.empty:
        return pd.DataFrame()
    pivot = (df.groupby(["dia", "zonal2"]).size()
               .reset_index(name="n")
               .pivot(index="dia", columns="zonal2", values="n")
               .fillna(0).astype(int))
    pivot.index = pivot.index.astype(int)
    pivot = pivot.reindex(range(1, DIA_CORTE + 1), fill_value=0)
    pivot.columns.name = None
    return pivot


def _vdd_unicos_mes_rus(periodo):
    """VDDs unicos del periodo completo para RUS."""
    df = pd.read_excel(BASE_DATOS_PATH, sheet_name="RUS",
                       usecols=["PERIODO", "Fecha_Registro", "zonal2", "DNI_VENDEDOR"])
    df = df[df["PERIODO"] == periodo].copy()
    df["dia"] = pd.to_datetime(df["Fecha_Registro"], errors="coerce").dt.day
    df = df[df["dia"].between(1, DIA_CORTE)]
    df = _filtrar_zonales(df)
    if df.empty:
        return pd.Series(dtype=int)
    return df.groupby("zonal2")["DNI_VENDEDOR"].nunique()


def _pivot_con():
    df = pd.read_excel(BASE_DATOS_PATH, sheet_name="CON")
    df["fecha_registro"] = pd.to_datetime(df["fecha_registro"], errors="coerce")
    df["dia"] = df["fecha_registro"].dt.day
    df = df[df["dia"].between(1, DIA_CORTE)].copy()
    df["ZONAL"] = df["ZONAL"].str.strip().str.upper()
    df = _filtrar_zonales(df, col="ZONAL")

    resultado = {}
    for periodo, clave in [("2026-04", "abr"), ("2026-05", "may")]:
        sub = df[df["PERIODO"] == periodo]
        if sub.empty:
            resultado[clave] = (pd.DataFrame(), pd.DataFrame(), pd.DataFrame())
            continue
        total = (sub.groupby(["dia", "ZONAL"])["CON"].sum().reset_index()
                    .pivot(index="dia", columns="ZONAL", values="CON")
                    .fillna(0).astype(int)
                    .reindex(range(1, DIA_CORTE + 1), fill_value=0))
        total.columns.name = None
        vend_u = (sub.groupby(["dia", "ZONAL"])["documento_v"].nunique()
                     .reset_index(name="n_vend")
                     .pivot(index="dia", columns="ZONAL", values="n_vend")
                     .fillna(0).astype(int)
                     .reindex(range(1, DIA_CORTE + 1), fill_value=0))
        vend_u.columns.name = None
        mrg = (sub.groupby(["dia", "ZONAL"])["CON"].sum().reset_index()
                  .merge(sub.groupby(["dia", "ZONAL"])["documento_v"]
                            .nunique().reset_index(name="n_vend"),
                         on=["dia", "ZONAL"]))
        mrg["ratio"] = (mrg["CON"] / mrg["n_vend"]).round(2)
        ratio = (mrg.pivot(index="dia", columns="ZONAL", values="ratio")
                    .reindex(range(1, DIA_CORTE + 1)))
        ratio = ratio.apply(pd.to_numeric, errors="coerce")
        ratio.columns.name = None
        resultado[clave] = (total, ratio, vend_u)
    return resultado


def _cargar_datos():
    print("    Cargando RT...")
    rt_abr = _pivot_rt_alt("RT",    "2026-04", "Fecha_Registro")
    rt_may = _pivot_rt_alt("RT",    "2026-05", "Fecha_Registro")
    print("    Cargando ALTAS...")
    alt_abr = _pivot_rt_alt("ALTAS", "2026-04", "Fecha_Alta")
    alt_may = _pivot_rt_alt("ALTAS", "2026-05", "Fecha_Alta")
    print("    Cargando RUS...")
    rus_abr = _pivot_rus("2026-04")
    rus_may = _pivot_rus("2026-05")
    print("    Cargando VDDs unicos por dia...")
    vdd_rt_abr  = _pivot_vdd_unicos("RT",    "2026-04", "Fecha_Registro")
    vdd_rt_may  = _pivot_vdd_unicos("RT",    "2026-05", "Fecha_Registro")
    vdd_alt_abr = _pivot_vdd_unicos("ALTAS", "2026-04", "Fecha_Alta")
    vdd_alt_may = _pivot_vdd_unicos("ALTAS", "2026-05", "Fecha_Alta")
    print("    Cargando VDDs unicos del mes completo...")
    vdd_rt_abr_mes  = _vdd_unicos_mes("RT",    "2026-04", "Fecha_Registro")
    vdd_rt_may_mes  = _vdd_unicos_mes("RT",    "2026-05", "Fecha_Registro")
    vdd_alt_abr_mes = _vdd_unicos_mes("ALTAS", "2026-04", "Fecha_Alta")
    vdd_alt_may_mes = _vdd_unicos_mes("ALTAS", "2026-05", "Fecha_Alta")
    vdd_rus_abr_mes = _vdd_unicos_mes_rus("2026-04")
    vdd_rus_may_mes = _vdd_unicos_mes_rus("2026-05")
    print("    Cargando CON...")
    con = _pivot_con()
    con_total_abr, con_ratio_abr, con_vdd_abr = con["abr"]
    con_total_may, con_ratio_may, con_vdd_may  = con["may"]
    con_vdd_abr_mes = _vdd_unicos_mes_con("2026-04")
    con_vdd_may_mes = _vdd_unicos_mes_con("2026-05")
    return dict(
        rt_abr=rt_abr, rt_may=rt_may,
        alt_abr=alt_abr, alt_may=alt_may,
        rus_abr=rus_abr, rus_may=rus_may,
        vdd_rt_abr=vdd_rt_abr, vdd_rt_may=vdd_rt_may,
        vdd_alt_abr=vdd_alt_abr, vdd_alt_may=vdd_alt_may,
        vdd_rt_abr_mes=vdd_rt_abr_mes, vdd_rt_may_mes=vdd_rt_may_mes,
        vdd_alt_abr_mes=vdd_alt_abr_mes, vdd_alt_may_mes=vdd_alt_may_mes,
        vdd_rus_abr_mes=vdd_rus_abr_mes, vdd_rus_may_mes=vdd_rus_may_mes,
        con_total_abr=con_total_abr, con_total_may=con_total_may,
        con_ratio_abr=con_ratio_abr, con_ratio_may=con_ratio_may,
        con_vdd_abr=con_vdd_abr, con_vdd_may=con_vdd_may,
        con_vdd_abr_mes=con_vdd_abr_mes, con_vdd_may_mes=con_vdd_may_mes,
    )


def _zonales_presentes(datos):
    zs = set()
    for k, df in datos.items():
        if isinstance(df, pd.DataFrame) and not df.empty:
            zs.update(df.columns.tolist())
        elif isinstance(df, pd.Series) and not df.empty:
            zs.update(df.index.tolist())
    # Filtrar explicitamente las excluidas
    zs -= ZONAL_EXCLUIR
    ordered = [z for z in ZONAL_FULL if z in zs]
    ordered += sorted(z for z in zs if z not in ordered)
    return ordered


# ── Agregaciones ──────────────────────────────────────────────────────────────
def _agg_mes(df, is_ratio=False):
    if df is None or df.empty:
        return pd.Series(dtype=float)
    df = df.apply(pd.to_numeric, errors="coerce")
    if is_ratio:
        return df.mean().round(2)
    return df.sum().fillna(0).astype(int)


def _agg_semanas(df, anio, mes, is_ratio=False):
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy().apply(pd.to_numeric, errors="coerce")
    d["_s"] = [_sem_label(anio, mes, int(i)) for i in d.index]
    cols = [c for c in d.columns if not c.startswith("_")]
    if is_ratio:
        res = d.groupby("_s")[cols].mean().round(2)
    else:
        res = d.groupby("_s")[cols].sum().fillna(0).astype(int)
    sem_order = sorted(res.index, key=lambda s: int(s[1:]))
    return res.loc[sem_order]


# ══════════════════════════════════════════════════════════════════════════════
# HOJA DATA
# ══════════════════════════════════════════════════════════════════════════════

def _escribir_hoja_data(ws, datos, zonales):
    """
    Layout:
      Cols A-F  (left):  DATE | FECHA | KPI | ZONAL | ZONAL_ABREV | VALOR
                         Una fila por (dia, KPI, zonal) para cada periodo
      Cols K-O  (right): SEMANA | KPI | ZONAL | ZONAL_ABREV | VALOR
                         Una fila por (semana, KPI, zonal) para cada periodo
    """
    # Cabeceras
    headers_left  = ["DATE", "FECHA", "KPI", "ZONAL", "ZONAL_ABREV", "VALOR"]
    headers_right = ["SEMANA", "KPI", "ZONAL", "ZONAL_ABREV", "VALOR"]
    for ci, h in enumerate(headers_left, 1):
        cs(ws, 1, ci, h, bold=True, bg=C["sub_hdr"], sz=10)
    for ci, h in enumerate(headers_right, 11):
        cs(ws, 1, ci, h, bold=True, bg=C["sub_hdr"], sz=10)

    # ── Construir datos left section ──────────────────────────────────────────
    # Mapeo KPI -> (df_abr, df_may, is_ratio)
    kpi_map = {
        KPI_CON_RATIO: (datos["con_ratio_abr"], datos["con_ratio_may"], True),
        KPI_VDDS_RT:   (datos["vdd_rt_abr"],    datos["vdd_rt_may"],    False),
        KPI_VDDS_ALT:  (datos["vdd_alt_abr"],   datos["vdd_alt_may"],   False),
        KPI_VDDS_CON:  (datos["con_vdd_abr"],   datos["con_vdd_may"],   False),
    }

    # Fechas de inicio de cada periodo para la seccion left
    # Usamos todos los dias 1-DIA_CORTE de cada periodo
    left_rows = []
    for kpi in KPIS_ORDER:
        df_abr, df_may, _ = kpi_map[kpi]
        for periodo_lbl, df, anio, mes in [
            ("ABR", df_abr, ANIO_ABR, MES_ABR),
            ("MAY", df_may, ANIO_MAY, MES_MAY),
        ]:
            for dia in range(1, DIA_CORTE + 1):
                for z in zonales:
                    valor = None
                    if df is not None and not df.empty and dia in df.index and z in df.columns:
                        valor = _safe(df.loc[dia, z])
                    left_rows.append({
                        "fecha": date(anio, mes, dia),
                        "kpi":   kpi,
                        "zonal": z,
                        "abrev": _abrev(z),
                        "valor": valor,
                    })

    # Construir datos right section
    right_rows = []
    for kpi in KPIS_ORDER:
        df_abr, df_may, is_ratio = kpi_map[kpi]
        for periodo_lbl, df, anio, mes in [
            ("ABR", df_abr, ANIO_ABR, MES_ABR),
            ("MAY", df_may, ANIO_MAY, MES_MAY),
        ]:
            sem_df = _agg_semanas(df, anio, mes, is_ratio)
            for sem in (sem_df.index.tolist() if not sem_df.empty else []):
                for z in zonales:
                    valor = None
                    if z in sem_df.columns:
                        valor = _safe(sem_df.loc[sem, z])
                    right_rows.append({
                        "semana": f"{periodo_lbl} {sem}",
                        "kpi":   kpi,
                        "zonal": z,
                        "abrev": _abrev(z),
                        "valor": valor,
                    })

    # Escribir left section
    for ri, row in enumerate(left_rows, 2):
        fecha = row["fecha"]
        # Col A: etiqueta fecha en formato "DD-MMM" con mayuscula (ej: "03-Abr")
        # Mismo formato que las etiquetas del RESUMEN para que los SUMIFS coincidan
        mes_abrev = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
                     7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
        fecha_lbl = f"{fecha.day:02d}-{mes_abrev[fecha.month]}"
        c = ws.cell(row=ri, column=1, value=fecha_lbl)
        c.font = Font(name=FONT_NAME, size=FONT_SIZE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        # Col B: fecha directa (puede repetirse para misma fecha en distintas zonales)
        c2 = ws.cell(row=ri, column=2, value=fecha)
        c2.font = Font(name=FONT_NAME, size=FONT_SIZE)
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.number_format = "YYYY-MM-DD"
        # Cols C-F
        ws.cell(row=ri, column=3, value=row["kpi"]).font = Font(name=FONT_NAME, size=FONT_SIZE)
        ws.cell(row=ri, column=4, value=row["zonal"]).font = Font(name=FONT_NAME, size=FONT_SIZE)
        ws.cell(row=ri, column=5, value=row["abrev"]).font = Font(name=FONT_NAME, size=FONT_SIZE)
        c6 = ws.cell(row=ri, column=6, value=row["valor"] if row["valor"] is not None else 0)
        c6.font = Font(name=FONT_NAME, size=FONT_SIZE)
        for ci in range(1, 7):
            ws.cell(row=ri, column=ci).alignment = Alignment(horizontal="center", vertical="center")

    # Escribir right section
    for ri, row in enumerate(right_rows, 2):
        ws.cell(row=ri, column=11, value=row["semana"]).font = Font(name=FONT_NAME, size=FONT_SIZE)
        ws.cell(row=ri, column=12, value=row["kpi"]).font = Font(name=FONT_NAME, size=FONT_SIZE)
        ws.cell(row=ri, column=13, value=row["zonal"]).font = Font(name=FONT_NAME, size=FONT_SIZE)
        ws.cell(row=ri, column=14, value=row["abrev"]).font = Font(name=FONT_NAME, size=FONT_SIZE)
        c15 = ws.cell(row=ri, column=15, value=row["valor"] if row["valor"] is not None else 0)
        c15.font = Font(name=FONT_NAME, size=FONT_SIZE)
        for ci in range(11, 16):
            ws.cell(row=ri, column=ci).alignment = Alignment(horizontal="center", vertical="center")

    # Anchos
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["K"].width = 10
    ws.column_dimensions["L"].width = 12
    ws.column_dimensions["M"].width = 16
    ws.column_dimensions["N"].width = 10
    ws.column_dimensions["O"].width = 10

    return len(left_rows) + 1  # ultima fila usada


# ══════════════════════════════════════════════════════════════════════════════
# HOJA RESUMEN — layout del ejemplo
# ══════════════════════════════════════════════════════════════════════════════

def _escribir_hoja_resumen(ws, datos, zonales):
    """
    Layout (basado en ARCHIVO_EJEMPLO_DRILL_DOWN.xlsx):

    3 meta-bloques horizontales: ABRIL (A-H) | MAYO (J-Q) | %VARIACION (S-Z)
    Col I, R = gap

    Row 1: Info | DEL 1 AL 17 (mergeado en cada bloque)
    Row 2: vacia
    Row 3: cabeceras zonales abreviadas
    Row 4: CON Ratio
    Row 5: RT
    Row 6: ALTAS
    Row 7: %Conv RT->ALT
    Row 8: %Conv CON->ALT

    Row 9-10: vacia
    Row 11: titulos de 4 sub-bloques
    Row 12: MES + zonales
    Row 13: ABRIL
    Row 14: MAYO
    Row 15: DIF
    Row 16: %DIF
    Row 17: vacia
    Row 18: SEM + zonales
    Row 19+: datos SEM (uno por semana equivalent)
    Row N+2: vacia
    Row N+3: DIA + zonales
    Row N+4+: datos DIA (uno por par equivalente)

    4 sub-bloques: cols A-H, J-Q, S-Z, AB-AI (con gap col I, R, AA entre bloques de 9 cols)
    Cada sub-bloque = 1 col label + 7 cols zonales
    """
    # Zonales a usar (solo los que tienen datos)
    zs = set()
    for df in datos.values():
        if isinstance(df, pd.DataFrame) and not df.empty:
            zs.update(df.columns.tolist())
    zonales_res = [z for z in ZONAL_FULL if z in zs]
    N_Z = len(zonales_res)

    # Layout de los 3 meta-bloques (filas 1-8): ABRIL | MAYO | %VARIACION
    # Cada meta-bloque = 1 col label + N_Z cols zonales, separados por 1 gap col
    # Ejemplo con 7 zonales: A(1)-H(8) | gap I(9) | J(10)-Q(17) | gap R(18) | S(19)-Z(26)
    ANCHO_META = 1 + N_Z  # col label + N_Z zonales
    GAP_META   = 1

    def col_meta(bi):   # bi=0 ABRIL, bi=1 MAYO, bi=2 %VARIACION
        return 1 + bi * (ANCHO_META + GAP_META)

    def col_z_meta(bi, zi):
        return col_meta(bi) + 1 + zi

    total_meta_cols = col_meta(2) + ANCHO_META  # hasta ultimo col de %VARIACION

    # ── Fila 1: INFO ──────────────────────────────────────────────────────────
    info_txt = f"Info  |  DEL  1  AL  {DIA_CORTE}"
    for bi, lbl in enumerate(["ABRIL", "MAYO", "%VARIACION"]):
        c0 = col_meta(bi)
        cs(ws, 1, c0, info_txt, bold=True, bg=C["hdr_dark"], fc=C["white"],
           align="center", sz=11)
        for cc in range(c0 + 1, c0 + ANCHO_META):
            cs(ws, 1, cc, None, bg=C["hdr_dark"])
        if c0 + ANCHO_META - 1 > c0:
            ws.merge_cells(start_row=1, start_column=c0,
                           end_row=1, end_column=c0 + ANCHO_META - 1)
        if bi < 2:
            ws.cell(row=1, column=c0 + ANCHO_META).value = None
    ws.row_dimensions[1].height = 20

    # ── Fila 2: vacia ─────────────────────────────────────────────────────────
    ws.row_dimensions[2].height = 6

    # ── Fila 3: cabeceras zonales ─────────────────────────────────────────────
    meta_lbls = ["ABRIL", "MAYO", "%VARIACION"]
    for bi, lbl in enumerate(meta_lbls):
        c0 = col_meta(bi)
        cs(ws, 3, c0, lbl, bold=True, bg=C["mes_hdr"], sz=9)
        for zi, z in enumerate(zonales_res):
            cs(ws, 3, col_z_meta(bi, zi), _abrev(z), bold=True, bg=C["sub_hdr"], sz=9)
        if bi < 2:
            ws.cell(row=3, column=c0 + ANCHO_META).value = None
    ws.row_dimensions[3].height = 14

    # ── Filas 4-8: funnel de metricas ─────────────────────────────────────────
    # Calcular valores para ABRIL y MAYO
    def _get_s(key, is_ratio=False):
        df = datos.get(key)
        return _agg_mes(df, is_ratio) if df is not None else pd.Series(dtype=float)

    con_ratio_abr_s = _get_s("con_ratio_abr", True)
    con_ratio_may_s = _get_s("con_ratio_may", True)
    con_total_abr_s = _get_s("con_total_abr")
    con_total_may_s = _get_s("con_total_may")
    rt_abr_s  = _get_s("rt_abr")
    rt_may_s  = _get_s("rt_may")
    alt_abr_s = _get_s("alt_abr")
    alt_may_s = _get_s("alt_may")
    rus_abr_s = _get_s("rus_abr")
    rus_may_s = _get_s("rus_may")

    def _conv(num_s, den_s, z):
        n = _safe(num_s.get(z) if hasattr(num_s, 'get') else None)
        d = _safe(den_s.get(z) if hasattr(den_s, 'get') else None)
        return round(n / d * 100, 1) if (n and d) else None

    funnel_def = [
        # (label, ser_abr, ser_may, num_fmt)
        ("CON Ratio",        con_ratio_abr_s, con_ratio_may_s, "0.00"),
        ("RT",               rt_abr_s,        rt_may_s,        "#,##0"),
        ("RUS",              rus_abr_s,        rus_may_s,       "#,##0"),
        ("ALTAS",            alt_abr_s,       alt_may_s,       "#,##0"),
        ("%Conv RT->RUS",    None,            None,            '0.0"%"'),
        ("%Conv RUS->ALTA",  None,            None,            '0.0"%"'),
        ("%Conv CON->ALTA",  None,            None,            '0.0"%"'),
    ]

    row_bgs_funnel = [
        C["con_row"], C["rt_row"], C["rus_row"], C["alt_row"],
        C["dia_hdr"], C["dia_hdr"], C["dia_hdr"],
    ]

    def _val_funnel(fi, z, usar_abr):
        """Devuelve el valor numerico para la fila fi del funnel, abril o mayo."""
        if fi == 4:   # %Conv RT->RUS
            return _conv(rus_abr_s if usar_abr else rus_may_s,
                         rt_abr_s  if usar_abr else rt_may_s, z)
        elif fi == 5:  # %Conv RUS->ALTA
            return _conv(alt_abr_s if usar_abr else alt_may_s,
                         rus_abr_s if usar_abr else rus_may_s, z)
        elif fi == 6:  # %Conv CON->ALTA
            return _conv(alt_abr_s if usar_abr else alt_may_s,
                         con_total_abr_s if usar_abr else con_total_may_s, z)
        else:
            ser = funnel_def[fi][1] if usar_abr else funnel_def[fi][2]
            return _safe(ser.get(z)) if ser is not None else None

    for fi, (lbl, ser_a, ser_m, nfmt) in enumerate(funnel_def):
        row = 4 + fi
        bg = row_bgs_funnel[fi]
        # ABRIL block (bi=0)
        c0 = col_meta(0)
        cs(ws, row, c0, lbl, bold=True, bg=bg, align="left", sz=9)
        for zi, z in enumerate(zonales_res):
            v = _val_funnel(fi, z, usar_abr=True)
            val = round(v, 2) if (v is not None and nfmt == "0.00") else (int(round(v)) if (v is not None and "#" in nfmt) else v)
            cs(ws, row, col_z_meta(0, zi),
               val if val is not None else "-",
               bg=bg, num_fmt=nfmt if val is not None and val != "-" else None, sz=9)
        ws.cell(row=row, column=col_meta(0) + ANCHO_META).value = None

        # MAYO block (bi=1)
        c1 = col_meta(1)
        cs(ws, row, c1, lbl, bold=True, bg=bg, align="left", sz=9)
        for zi, z in enumerate(zonales_res):
            v = _val_funnel(fi, z, usar_abr=False)
            val = round(v, 2) if (v is not None and nfmt == "0.00") else (int(round(v)) if (v is not None and "#" in nfmt) else v)
            cs(ws, row, col_z_meta(1, zi),
               val if val is not None else "-",
               bg=bg, num_fmt=nfmt if val is not None and val != "-" else None, sz=9)
        ws.cell(row=row, column=col_meta(1) + ANCHO_META).value = None

        # %VARIACION block (bi=2): formulas Excel =(Kx-Bx)/Bx
        c2 = col_meta(2)
        cs(ws, row, c2, lbl, bold=True, bg=bg, align="left", sz=9)
        for zi in range(N_Z):
            col_abr = get_column_letter(col_z_meta(0, zi))
            col_may = get_column_letter(col_z_meta(1, zi))
            formula = f"=IFERROR(({col_may}{row}-{col_abr}{row})/{col_abr}{row},\"-\")"
            c = ws.cell(row=row, column=col_z_meta(2, zi), value=formula)
            c.font  = Font(name=FONT_NAME, size=FONT_SIZE - 2)
            c.fill  = PatternFill("solid", start_color=bg)
            c.number_format = '0.0%'
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = _border()
        ws.row_dimensions[row].height = 13

    # ── Filas 8-10: altura fija + separador entre tablas ─────────────────────
    # Filas 8-10 son las ultimas 3 metricas del funnel (fi=4,5,6 => rows 8,9,10)
    # y la fila 11 es el inicio de los sub-bloques
    for r in range(8, 11):
        ws.row_dimensions[r].height = 15
    # Fila de espacio entre el funnel y los sub-bloques
    # El funnel ocupa filas 4..4+len(funnel_def)-1, luego insertamos un separador
    row_sep = 4 + len(funnel_def)   # fila justo despues del ultimo KPI
    ws.row_dimensions[row_sep].height = 8

    # ── 4 sub-bloques para MES/SEM/DIA ───────────────────────────────────────
    # Layout: 4 bloques horizontales, cada uno = 1 col label + N_Z cols zonales + 1 gap
    # Bloque 0 empieza en col 1, bloque 1 en col ANCHO_META+GAP+1, etc.
    # Con N_Z=7: bloque 0=A-H, gap I, bloque 1=J-Q, gap R, bloque 2=S-Z, gap AA, bloque 3=AB-AI
    ANCHO_SUB = 1 + N_Z
    GAP_SUB   = 1

    def col_sub(bi):
        return 1 + bi * (ANCHO_SUB + GAP_SUB)

    def col_z_sub(bi, zi):
        return col_sub(bi) + 1 + zi

    # Titulos de los 4 sub-bloques
    sub_titulos = [
        "RATIO CONSULTAS",
        "VDDS DIFERENTES CONSULTANDO",
        "VDDS DIFERENTES CON VENTA (RT)",
        "VDDS DIFERENTES CON ALTA",
    ]
    sub_kpis = [KPI_CON_RATIO, KPI_VDDS_CON, KPI_VDDS_RT, KPI_VDDS_ALT]
    sub_hdr_colors = [C["con_hdr"], C["con_hdr"], C["rt_hdr"], C["alt_hdr"]]
    sub_is_ratio = [True, False, False, False]
    sub_nfmts    = ["0.00", "#,##0", "#,##0", "#,##0"]
    # Datos por sub-bloque: pivot diario (para SEM/DIA via DATA sheet)
    sub_data_abr = [
        datos["con_ratio_abr"],
        datos["con_vdd_abr"],
        datos["vdd_rt_abr"],
        datos["vdd_alt_abr"],
    ]
    sub_data_may = [
        datos["con_ratio_may"],
        datos["con_vdd_may"],
        datos["vdd_rt_may"],
        datos["vdd_alt_may"],
    ]
    # Series MES correctas (VDDs unicos del periodo, no suma de dias)
    sub_mes_abr = [
        _agg_mes(datos["con_ratio_abr"], True),   # CON RATIO: promedio de dias ok
        datos["con_vdd_abr_mes"],                  # VDDs CON: unicos del mes
        datos["vdd_rt_abr_mes"],                   # VDDs RT:  unicos del mes
        datos["vdd_alt_abr_mes"],                  # VDDs ALT: unicos del mes
    ]
    sub_mes_may = [
        _agg_mes(datos["con_ratio_may"], True),
        datos["con_vdd_may_mes"],
        datos["vdd_rt_may_mes"],
        datos["vdd_alt_may_mes"],
    ]

    # row_sep es la fila de espacio entre funnel y sub-bloques (ya definida arriba)
    ROW_SUB_HDR = row_sep + 1   # fila de titulos de los 4 sub-bloques

    # ── Titulos de sub-bloques ────────────────────────────────────────────────
    for bi, titulo in enumerate(sub_titulos):
        c0 = col_sub(bi)
        cs(ws, ROW_SUB_HDR, c0, titulo, bold=True, bg=sub_hdr_colors[bi], fc=C["white"],
           align="left", sz=9)
        for cc in range(c0 + 1, c0 + ANCHO_SUB):
            cs(ws, ROW_SUB_HDR, cc, None, bg=sub_hdr_colors[bi])
        if c0 + ANCHO_SUB - 1 > c0:
            ws.merge_cells(start_row=ROW_SUB_HDR, start_column=c0,
                           end_row=ROW_SUB_HDR, end_column=c0 + ANCHO_SUB - 1)
        if bi < 3:
            ws.cell(row=ROW_SUB_HDR, column=c0 + ANCHO_SUB).value = None
    ws.row_dimensions[ROW_SUB_HDR].height = 16

    # ── MES + zonales ─────────────────────────────────────────────────────────
    ROW_MES_HDR = ROW_SUB_HDR + 1
    for bi in range(4):
        c0 = col_sub(bi)
        cs(ws, ROW_MES_HDR, c0, "MES", bold=True, bg=C["mes_hdr"], sz=9)
        for zi, z in enumerate(zonales_res):
            cs(ws, ROW_MES_HDR, col_z_sub(bi, zi), _abrev(z), bold=True, bg=C["sub_hdr"], sz=9)
        if bi < 3:
            ws.cell(row=ROW_MES_HDR, column=c0 + ANCHO_SUB).value = None
    ws.row_dimensions[ROW_MES_HDR].height = 14

    # ── ABRIL / MAYO / DIF / %DIF ─────────────────────────────────────────────
    mes_rows = [("ABRIL", "abr"), ("MAYO", "may"), ("DIF", None), ("%DIF", None)]
    mes_bgs  = [C["abr"], C["may"], C["dlt"], C["pct"]]
    mes_fcs  = ["000000", C["white"], "000000", "000000"]

    for mi, ((lbl, key_suf), bg, fc) in enumerate(zip(mes_rows, mes_bgs, mes_fcs)):
        row = ROW_MES_HDR + 1 + mi
        for bi in range(4):
            c0 = col_sub(bi)
            cs(ws, row, c0, lbl, bold=True, bg=bg, fc=fc, align="left", sz=9)
            is_r = sub_is_ratio[bi]
            nfmt = sub_nfmts[bi]
            s_a  = sub_mes_abr[bi]
            s_m  = sub_mes_may[bi]
            for zi, z in enumerate(zonales_res):
                va = _safe(s_a.get(z))
                vm = _safe(s_m.get(z))
                if key_suf == "abr":
                    val = round(va, 2) if (va is not None and is_r) else (int(round(va)) if va is not None else None)
                elif key_suf == "may":
                    val = round(vm, 2) if (vm is not None and is_r) else (int(round(vm)) if vm is not None else None)
                elif lbl == "DIF":
                    val = round(vm - va, 2) if (va is not None and vm is not None) else None
                else:  # %DIF
                    val = round((vm - va) / abs(va) * 100, 1) if (va is not None and vm is not None and va != 0) else None
                dfc = (C["pos"] if (val is not None and lbl in ("DIF", "%DIF") and val >= 0) else
                       C["neg"] if (val is not None and lbl in ("DIF", "%DIF") and val < 0) else fc)
                use_nfmt = ('0.0"%"' if lbl == "%DIF" else nfmt) if val is not None else None
                cs(ws, row, col_z_sub(bi, zi),
                   val if val is not None else "-",
                   bg=bg, fc=dfc, num_fmt=use_nfmt, sz=9)
            if bi < 3:
                ws.cell(row=row, column=c0 + ANCHO_SUB).value = None
        ws.row_dimensions[row].height = 13

    # fila de espacio antes de SEM
    ws.row_dimensions[ROW_MES_HDR + 5].height = 8

    # ── SEM + zonales ─────────────────────────────────────────────────────────
    ROW_SEM_HDR = ROW_MES_HDR + 6
    for bi in range(4):
        c0 = col_sub(bi)
        cs(ws, ROW_SEM_HDR, c0, "SEM", bold=True, bg=C["sem_hdr"], sz=9)
        for zi, z in enumerate(zonales_res):
            cs(ws, ROW_SEM_HDR, col_z_sub(bi, zi), _abrev(z), bold=True, bg=C["sub_hdr"], sz=9)
        if bi < 3:
            ws.cell(row=ROW_SEM_HDR, column=c0 + ANCHO_SUB).value = None
    ws.row_dimensions[ROW_SEM_HDR].height = 14

    # ── Semanas equivalentes (ABR Sx | MAY Sx) ────────────────────────────────
    # Calcular semanas de ambos meses
    all_sems = set()
    for bi in range(4):
        for df, anio, mes in [(sub_data_abr[bi], ANIO_ABR, MES_ABR),
                               (sub_data_may[bi], ANIO_MAY, MES_MAY)]:
            sem_df = _agg_semanas(df, anio, mes, sub_is_ratio[bi])
            all_sems.update(sem_df.index.tolist() if not sem_df.empty else [])
    all_sems = sorted(all_sems, key=lambda s: int(s[1:]))

    cur = ROW_SEM_HDR + 1
    row_bgs2 = [C["info_hdr"], C["white"]]

    for si, sem in enumerate(all_sems):
        bg_s = row_bgs2[si % 2]
        lbl_sem = f"ABR {sem} | MAY {sem}"

        # Col A del bloque 0: label
        cs(ws, cur, col_sub(0), lbl_sem, bold=False, bg=bg_s, align="left", sz=9)
        # Cols B-H del bloque 0: SUMIFS formulas para cada KPI segun sub-bloque
        # El ejemplo muestra formulas solo para el primer sub-bloque (RATIO CONSULTAS / CON RATIO)
        # Los demas bloques usan =A{row} para la etiqueta y formulas propias para los datos
        for zi in range(N_Z):
            col_letter = get_column_letter(col_z_sub(0, zi))
            row_hdr = ROW_SEM_HDR
            hdr_col_letter = get_column_letter(col_z_sub(0, zi))
            kpi_name = sub_kpis[0]  # CON RATIO para bloque 0
            formula = (
                f'=IFERROR((SUMIFS(DATA!$O:$O,DATA!$K:$K,RIGHT(RESUMEN!$A{cur},6),'
                f'DATA!$N:$N,{hdr_col_letter}${row_hdr},DATA!$L:$L,"{kpi_name}")'
                f'-SUMIFS(DATA!$O:$O,DATA!$K:$K,LEFT(RESUMEN!$A{cur},6),'
                f'DATA!$N:$N,{hdr_col_letter}${row_hdr},DATA!$L:$L,"{kpi_name}"))'
                f'/SUMIFS(DATA!$O:$O,DATA!$K:$K,LEFT(RESUMEN!$A{cur},6),'
                f'DATA!$N:$N,{hdr_col_letter}${row_hdr},DATA!$L:$L,"{kpi_name}"),"-")'
            )
            c = ws.cell(row=cur, column=col_z_sub(0, zi), value=formula)
            c.font = Font(name=FONT_NAME, size=FONT_SIZE - 2)
            c.fill = PatternFill("solid", start_color=bg_s)
            c.number_format = "0.0%"
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = _border()

        # Bloques 1-3: etiqueta = =A{cur}, datos = SUMIFS con su propio KPI
        for bi in range(1, 4):
            col_lbl = get_column_letter(col_sub(0))
            col_lbl_bi = col_sub(bi)
            # La etiqueta referencia el bloque anterior
            prev_lbl_col = get_column_letter(col_sub(bi - 1))
            c_lbl = ws.cell(row=cur, column=col_lbl_bi,
                            value=f"={prev_lbl_col}{cur}")
            c_lbl.font = Font(name=FONT_NAME, size=FONT_SIZE - 2)
            c_lbl.fill = PatternFill("solid", start_color=bg_s)
            c_lbl.alignment = Alignment(horizontal="left", vertical="center")
            c_lbl.border = _border()
            # Datos con SUMIFS para el KPI de este sub-bloque
            kpi_name = sub_kpis[bi]
            for zi in range(N_Z):
                hdr_col_letter = get_column_letter(col_z_sub(bi, zi))
                row_hdr = ROW_SEM_HDR
                formula = (
                    f'=IFERROR((SUMIFS(DATA!$O:$O,DATA!$K:$K,RIGHT(RESUMEN!${get_column_letter(col_sub(0))}{cur},6),'
                    f'DATA!$N:$N,{hdr_col_letter}${row_hdr},DATA!$L:$L,"{kpi_name}")'
                    f'-SUMIFS(DATA!$O:$O,DATA!$K:$K,LEFT(RESUMEN!${get_column_letter(col_sub(0))}{cur},6),'
                    f'DATA!$N:$N,{hdr_col_letter}${row_hdr},DATA!$L:$L,"{kpi_name}"))'
                    f'/SUMIFS(DATA!$O:$O,DATA!$K:$K,LEFT(RESUMEN!${get_column_letter(col_sub(0))}{cur},6),'
                    f'DATA!$N:$N,{hdr_col_letter}${row_hdr},DATA!$L:$L,"{kpi_name}"),"-")'
                )
                c = ws.cell(row=cur, column=col_z_sub(bi, zi), value=formula)
                c.font = Font(name=FONT_NAME, size=FONT_SIZE - 2)
                c.fill = PatternFill("solid", start_color=bg_s)
                c.number_format = "0.0%"
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = _border()
            if bi < 3:
                ws.cell(row=cur, column=col_sub(bi) + ANCHO_SUB).value = None

        ws.row_dimensions[cur].height = 13
        cur += 1

    # ── Espacio entre SEM y DIA ───────────────────────────────────────────────
    ws.row_dimensions[cur].height = 8
    cur += 1

    # ── DIA + zonales ─────────────────────────────────────────────────────────
    ROW_DIA_HDR = cur
    for bi in range(4):
        c0 = col_sub(bi)
        cs(ws, ROW_DIA_HDR, c0, "DIA", bold=True, bg=C["dia_hdr"], sz=9)
        for zi, z in enumerate(zonales_res):
            cs(ws, ROW_DIA_HDR, col_z_sub(bi, zi), _abrev(z), bold=True, bg=C["sub_hdr"], sz=9)
        if bi < 3:
            ws.cell(row=ROW_DIA_HDR, column=c0 + ANCHO_SUB).value = None
    ws.row_dimensions[ROW_DIA_HDR].height = 14
    cur += 1

    # ── Dias equivalentes ─────────────────────────────────────────────────────
    for di, (d_a, d_m) in enumerate(DIAS_EQUIV):
        bg_d = row_bgs2[di % 2]
        lbl_d = _lbl_dia(d_a, d_m)

        # Bloque 0: label + SUMIFS por CON RATIO
        cs(ws, cur, col_sub(0), lbl_d, bold=False, bg=bg_d, align="left", sz=9)
        kpi_name = sub_kpis[0]
        for zi in range(N_Z):
            hdr_col_letter = get_column_letter(col_z_sub(0, zi))
            row_hdr = ROW_DIA_HDR
            formula = (
                f'=IFERROR((SUMIFS(DATA!$F:$F,DATA!$A:$A,RIGHT(RESUMEN!$A{cur},6),'
                f'DATA!$E:$E,{hdr_col_letter}${row_hdr},DATA!$C:$C,"{kpi_name}")'
                f'-SUMIFS(DATA!$F:$F,DATA!$A:$A,LEFT(RESUMEN!$A{cur},6),'
                f'DATA!$E:$E,{hdr_col_letter}${row_hdr},DATA!$C:$C,"{kpi_name}"))'
                f'/SUMIFS(DATA!$F:$F,DATA!$A:$A,LEFT(RESUMEN!$A{cur},6),'
                f'DATA!$E:$E,{hdr_col_letter}${row_hdr},DATA!$C:$C,"{kpi_name}"),"-")'
            )
            c = ws.cell(row=cur, column=col_z_sub(0, zi), value=formula)
            c.font = Font(name=FONT_NAME, size=FONT_SIZE - 2)
            c.fill = PatternFill("solid", start_color=bg_d)
            c.number_format = "0.0%"
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = _border()

        # Bloques 1-3: label referencia bloque anterior, SUMIFS con su KPI
        for bi in range(1, 4):
            prev_lbl_col = get_column_letter(col_sub(bi - 1))
            c_lbl = ws.cell(row=cur, column=col_sub(bi),
                            value=f"={prev_lbl_col}{cur}")
            c_lbl.font = Font(name=FONT_NAME, size=FONT_SIZE - 2)
            c_lbl.fill = PatternFill("solid", start_color=bg_d)
            c_lbl.alignment = Alignment(horizontal="left", vertical="center")
            c_lbl.border = _border()
            kpi_name = sub_kpis[bi]
            for zi in range(N_Z):
                hdr_col_letter = get_column_letter(col_z_sub(bi, zi))
                row_hdr = ROW_DIA_HDR
                lbl_col_a = get_column_letter(col_sub(0))
                formula = (
                    f'=IFERROR((SUMIFS(DATA!$F:$F,DATA!$A:$A,RIGHT(RESUMEN!${lbl_col_a}{cur},6),'
                    f'DATA!$E:$E,{hdr_col_letter}${row_hdr},DATA!$C:$C,"{kpi_name}")'
                    f'-SUMIFS(DATA!$F:$F,DATA!$A:$A,LEFT(RESUMEN!${lbl_col_a}{cur},6),'
                    f'DATA!$E:$E,{hdr_col_letter}${row_hdr},DATA!$C:$C,"{kpi_name}"))'
                    f'/SUMIFS(DATA!$F:$F,DATA!$A:$A,LEFT(RESUMEN!${lbl_col_a}{cur},6),'
                    f'DATA!$E:$E,{hdr_col_letter}${row_hdr},DATA!$C:$C,"{kpi_name}"),"-")'
                )
                c = ws.cell(row=cur, column=col_z_sub(bi, zi), value=formula)
                c.font = Font(name=FONT_NAME, size=FONT_SIZE - 2)
                c.fill = PatternFill("solid", start_color=bg_d)
                c.number_format = "0.0%"
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = _border()
            if bi < 3:
                ws.cell(row=cur, column=col_sub(bi) + ANCHO_SUB).value = None

        ws.row_dimensions[cur].height = 13
        cur += 1

    # ── Anchos de columna ─────────────────────────────────────────────────────
    for bi in range(4):
        c0 = col_sub(bi)
        ws.column_dimensions[get_column_letter(c0)].width = 22
        for zi in range(N_Z):
            ws.column_dimensions[get_column_letter(col_z_sub(bi, zi))].width = 7
        if bi < 3:
            ws.column_dimensions[get_column_letter(c0 + ANCHO_SUB)].width = 2

    ws.freeze_panes = "B4"


# ══════════════════════════════════════════════════════════════════════════════
# HOJAS RT / ALT / CON
# ══════════════════════════════════════════════════════════════════════════════

def _escribir_hoja_metrica(ws, metrica, datos, zonales):
    """
    Layout (basado en ejemplo):
      Col A = etiqueta
      Por cada zonal: 4 cols (ABR|MAY|DIF|%DIF), separadas por 1 col gap

    Filas:
      1 = INFO header
      2 = zonal names (mergeado 4 cols cada uno)
      3 = RANGO / etiqueta metrica
      4 = MES / ABR|MAY|DIF|%DIF subcabecera
      5 = fila de datos MES
      6 = (vacia)
      7 = SEMANA header
      8 = SEM / ABR|MAY|DIF|%DIF subcabecera
      9..N = semanas
      N+1 = (vacia)
      N+2 = DIA header
      N+3 = DIA / ABR|MAY|DIF|%DIF subcabecera
      N+4.. = pares de dias
    """
    if metrica == "CON":
        df_abr  = datos["con_ratio_abr"]
        df_may  = datos["con_ratio_may"]
        is_rat  = True
        hdr_col = C["con_hdr"]
        titulo  = "CONSULTAS (ratio CON / vendedor)"
        num_fmt = "0.00"
        lbl_met = "CON ratio"
    elif metrica == "RT":
        df_abr  = datos["rt_abr"]
        df_may  = datos["rt_may"]
        is_rat  = False
        hdr_col = C["rt_hdr"]
        titulo  = "REGISTROS TOTALES"
        num_fmt = "#,##0"
        lbl_met = "RT"
    elif metrica == "RUS":
        df_abr  = datos["rus_abr"]
        df_may  = datos["rus_may"]
        is_rat  = False
        hdr_col = C["rus_hdr"]
        titulo  = "RUS - Registros Unicos de Servicio (metrica principal Movistar)"
        num_fmt = "#,##0"
        lbl_met = "RUS"
    else:
        df_abr  = datos["alt_abr"]
        df_may  = datos["alt_may"]
        is_rat  = False
        hdr_col = C["alt_hdr"]
        titulo  = "ALTAS"
        num_fmt = "#,##0"
        lbl_met = "ALTA"

    zs = set()
    for df in [df_abr, df_may]:
        if df is not None and not df.empty:
            zs.update(df.columns.tolist())
    zs_list = [z for z in zonales if z in zs]
    zs_list += sorted(z for z in zs if z not in zs_list and z != "SIN ZONAL")
    N = len(zs_list)

    GAP   = 1
    ANCHO = 4  # ABR|MAY|DIF|%DIF

    def col_z(zi):
        return 2 + zi * (ANCHO + GAP)

    n_total = col_z(N - 1) + ANCHO if N > 0 else 6

    def _merge_hdr(row, c0, c1, val, bg, bold=True, fc="000000", sz=None):
        cs(ws, row, c0, val, bold=bold, bg=bg, fc=fc, align="center", sz=sz or FONT_SIZE)
        for cc in range(c0 + 1, c1 + 1):
            cs(ws, row, cc, None, bg=bg)
        if c1 > c0:
            ws.merge_cells(start_row=row, start_column=c0,
                           end_row=row, end_column=c1)

    # ── Fila 1: INFO ──────────────────────────────────────────────────────────
    _merge_hdr(1, 1, n_total,
               f"{titulo}   |   Info: DEL 1 AL {DIA_CORTE}",
               hdr_col, fc=C["white"], sz=11)
    ws.row_dimensions[1].height = 22

    # ── Fila 2: nombres de zonales mergeados ──────────────────────────────────
    cs(ws, 2, 1, "", bg=C["sub_hdr"])
    for zi, z in enumerate(zs_list):
        c0 = col_z(zi)
        _merge_hdr(2, c0, c0 + ANCHO - 1, z, C["sub_hdr"], sz=9)
        if zi < N - 1:
            ws.cell(row=2, column=c0 + ANCHO).value = None
    ws.row_dimensions[2].height = 16

    # ── Fila 3: RANGO y etiqueta metrica ─────────────────────────────────────
    cs(ws, 3, 1, "RANGO", bold=True, bg=C["sub_hdr"], sz=9)
    for zi in range(N):
        c0 = col_z(zi)
        for off in range(ANCHO):
            cs(ws, 3, c0 + off, None, bg=C["sub_hdr"])
        if zi < N - 1:
            ws.cell(row=3, column=c0 + ANCHO).value = None
    ws.row_dimensions[3].height = 5

    # ── Fila 4: MES subcabecera ───────────────────────────────────────────────
    cs(ws, 4, 1, "MES", bold=True, bg=C["mes_hdr"], sz=9)
    for zi in range(N):
        c0 = col_z(zi)
        cs(ws, 4, c0,     "ABRIL", bold=True, bg=C["abr"], sz=8)
        cs(ws, 4, c0 + 1, "MAYO",  bold=True, bg=C["may"], fc=C["white"], sz=8)
        cs(ws, 4, c0 + 2, "DIF",   bold=True, bg=C["dlt"], sz=8, italic=True)
        cs(ws, 4, c0 + 3, "%DIF",  bold=True, bg=C["pct"], sz=8)
        if zi < N - 1:
            ws.cell(row=4, column=c0 + ANCHO).value = None
    ws.row_dimensions[4].height = 14

    def _fila_datos(row, etq, row_abr, row_may, bg):
        cs(ws, row, 1, etq, bold=False, bg=bg, align="left", sz=9)
        for zi, z in enumerate(zs_list):
            c0 = col_z(zi)
            va = _safe((row_abr or {}).get(z))
            vm = _safe((row_may or {}).get(z))
            va_d = (round(va, 2) if is_rat else int(round(va))) if va is not None else "-"
            vm_d = (round(vm, 2) if is_rat else int(round(vm))) if vm is not None else "-"
            dv   = round(vm - va, 2) if (va is not None and vm is not None) else None
            dpct = round((vm - va) / abs(va) * 100, 1) if (va and vm is not None) else None
            dfc  = C["pos"] if (dv is not None and dv >= 0) else C["neg"]
            cs(ws, row, c0,     va_d, bg=bg, num_fmt=num_fmt if va is not None else None, sz=9)
            cs(ws, row, c0 + 1, vm_d, bg=bg, num_fmt=num_fmt if vm is not None else None, sz=9)
            cs(ws, row, c0 + 2,
               (round(dv, 2) if is_rat else int(round(dv))) if dv is not None else "-",
               bg=bg, num_fmt=num_fmt if dv is not None else None,
               fc=dfc if dv is not None else "000000", italic=True, sz=9)
            cs(ws, row, c0 + 3,
               dpct if dpct is not None else "-",
               bg=bg, num_fmt='0.0"%"' if dpct is not None else None,
               fc=dfc if dpct is not None else "000000", sz=9)
            if zi < N - 1:
                ws.cell(row=row, column=c0 + ANCHO).value = None

    # ── Fila 5: datos MES ─────────────────────────────────────────────────────
    s_a = _agg_mes(df_abr, is_rat)
    s_m = _agg_mes(df_may, is_rat)
    rng_lbl = f"01/04-{DIA_CORTE:02d}/04  |  01/05-{DIA_CORTE:02d}/05"
    _fila_datos(5, rng_lbl, s_a.to_dict(), s_m.to_dict(), C["info_hdr"])
    ws.row_dimensions[5].height = 13

    cur = 6
    ws.row_dimensions[cur].height = 6
    cur += 1

    # ── SEMANA ────────────────────────────────────────────────────────────────
    _merge_hdr(cur, 1, n_total, "SEMANA", C["sem_hdr"], sz=9)
    ws.row_dimensions[cur].height = 13
    cur += 1

    sem_abr = _agg_semanas(df_abr, ANIO_ABR, MES_ABR, is_rat)
    sem_may = _agg_semanas(df_may, ANIO_MAY, MES_MAY, is_rat)
    all_sems = sorted(
        set(list(sem_abr.index) if not sem_abr.empty else []) |
        set(list(sem_may.index) if not sem_may.empty else []),
        key=lambda s: int(s[1:]))

    cs(ws, cur, 1, "SEM", bold=True, bg=C["sem_hdr"], sz=9)
    for zi in range(N):
        c0 = col_z(zi)
        cs(ws, cur, c0,     "ABR", bold=True, bg=C["abr"], sz=8)
        cs(ws, cur, c0 + 1, "MAY", bold=True, bg=C["may"], fc=C["white"], sz=8)
        cs(ws, cur, c0 + 2, "DIF", bold=True, bg=C["dlt"], sz=8, italic=True)
        cs(ws, cur, c0 + 3, "%DIF", bold=True, bg=C["pct"], sz=8)
        if zi < N - 1:
            ws.cell(row=cur, column=c0 + ANCHO).value = None
    ws.row_dimensions[cur].height = 14
    cur += 1

    row_bgs = [C["info_hdr"], C["white"]]
    for si, sem in enumerate(all_sems):
        bg_s = row_bgs[si % 2]
        ra = sem_abr.loc[sem].to_dict() if (not sem_abr.empty and sem in sem_abr.index) else {}
        rm = sem_may.loc[sem].to_dict() if (not sem_may.empty and sem in sem_may.index) else {}
        lbl = f"ABR {sem} | MAY {sem}"
        _fila_datos(cur, lbl, ra, rm, bg_s)
        ws.row_dimensions[cur].height = 13
        cur += 1

    ws.row_dimensions[cur].height = 6
    cur += 1

    # ── DIA ───────────────────────────────────────────────────────────────────
    _merge_hdr(cur, 1, n_total, "DIA", C["dia_hdr"], sz=9)
    ws.row_dimensions[cur].height = 13
    cur += 1

    cs(ws, cur, 1, "DIA", bold=True, bg=C["dia_hdr"], sz=9)
    for zi in range(N):
        c0 = col_z(zi)
        cs(ws, cur, c0,     "ABR", bold=True, bg=C["abr"], sz=8)
        cs(ws, cur, c0 + 1, "MAY", bold=True, bg=C["may"], fc=C["white"], sz=8)
        cs(ws, cur, c0 + 2, "DIF", bold=True, bg=C["dlt"], sz=8, italic=True)
        cs(ws, cur, c0 + 3, "%DIF", bold=True, bg=C["pct"], sz=8)
        if zi < N - 1:
            ws.cell(row=cur, column=c0 + ANCHO).value = None
    ws.row_dimensions[cur].height = 14
    cur += 1

    for di, (d_a, d_m) in enumerate(DIAS_EQUIV):
        bg_d = row_bgs[di % 2]
        ra = df_abr.loc[d_a].to_dict() if (df_abr is not None and not df_abr.empty and d_a in df_abr.index) else {}
        rm = df_may.loc[d_m].to_dict() if (df_may is not None and not df_may.empty and d_m in df_may.index) else {}
        _fila_datos(cur, _lbl_dia(d_a, d_m), ra, rm, bg_d)
        ws.row_dimensions[cur].height = 13
        cur += 1

    # ── Anchos ────────────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 26
    for zi in range(N):
        c0 = col_z(zi)
        for off in range(ANCHO):
            ws.column_dimensions[get_column_letter(c0 + off)].width = 9
        if zi < N - 1:
            ws.column_dimensions[get_column_letter(c0 + ANCHO)].width = 2
    ws.freeze_panes = "B4"


# ══════════════════════════════════════════════════════════════════════════════
# HOJA FUNNEL — semaforo de conversion para el jefe de proyecto
# ══════════════════════════════════════════════════════════════════════════════

def _escribir_hoja_funnel(ws, datos, zonales):
    """
    Tabla resumen del funnel completo CON → RT → RUS → ALTA.
    Para cada zonal muestra los volumenes de cada etapa y los % de conversion,
    con semaforo verde/amarillo/rojo comparando Mayo vs Abril.

    Semaforo de conversion (RUS->ALTA, RT->RUS, CON->ALTA):
      Verde    : mayo >= abril (mejora o igual)
      Amarillo : mayo entre 80%-99% de abril (leve caida)
      Rojo     : mayo < 80% de abril (caida significativa)

    Semaforo de volumenes (RT, RUS, ALTAS):
      Verde    : mayo >= abril
      Rojo     : mayo < abril
    """
    def _sem_bg_conv(v_may, v_abr):
        if v_may is None or v_abr is None or v_abr == 0:
            return C["dlt"]
        ratio = v_may / v_abr
        if ratio >= 1.0:
            return C["sem_verde"]
        if ratio >= 0.80:
            return C["sem_amarillo"]
        return C["sem_rojo"]

    def _sem_fc_conv(v_may, v_abr):
        if v_may is None or v_abr is None or v_abr == 0:
            return "000000"
        return C["sem_verde_fc"] if v_may >= v_abr else C["sem_rojo_fc"]

    def _sem_bg_vol(v_may, v_abr):
        if v_may is None or v_abr is None:
            return C["dlt"]
        return C["sem_verde"] if v_may >= v_abr else C["sem_rojo"]

    # Agregar series de totales del periodo
    def _get_s(key, is_ratio=False):
        df = datos.get(key)
        return _agg_mes(df, is_ratio) if df is not None else pd.Series(dtype=float)

    rt_a  = _get_s("rt_abr")
    rt_m  = _get_s("rt_may")
    rus_a = _get_s("rus_abr")
    rus_m = _get_s("rus_may")
    alt_a = _get_s("alt_abr")
    alt_m = _get_s("alt_may")
    con_a = _get_s("con_total_abr")
    con_m = _get_s("con_total_may")

    # Zonales con datos
    zs = set()
    for s in [rt_a, rt_m, rus_a, rus_m, alt_a, alt_m]:
        if not s.empty:
            zs.update(s.index.tolist())
    zs_list = [z for z in ZONAL_FULL if z in zs]
    N = len(zs_list)

    # Layout: col A = etiqueta fila, cols B..N+1 = zonales, col N+2 = TOTAL
    COL_LBL   = 1
    COL_TOTAL = 2 + N

    # ── Fila 1: titulo ────────────────────────────────────────────────────────
    titulo = f"FUNNEL DE CONVERSION   |   Del 1 al {DIA_CORTE} de cada mes"
    cs(ws, 1, COL_LBL, titulo, bold=True, bg=C["hdr_dark"], fc=C["white"], sz=12)
    for c in range(2, COL_TOTAL + 1):
        cs(ws, 1, c, None, bg=C["hdr_dark"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=COL_TOTAL)
    ws.row_dimensions[1].height = 22

    # ── Fila 2: cabecera zonales ───────────────────────────────────────────────
    cs(ws, 2, COL_LBL, "INDICADOR", bold=True, bg=C["sub_hdr"], sz=9)
    for zi, z in enumerate(zs_list):
        cs(ws, 2, 2 + zi, _abrev(z), bold=True, bg=C["sub_hdr"], sz=9)
    cs(ws, 2, COL_TOTAL, "TOTAL", bold=True, bg=C["sub_hdr"], sz=9)
    ws.row_dimensions[2].height = 14

    # ── Bloque de filas: una por indicador ────────────────────────────────────
    # Formato: (etiqueta, serie_abr, serie_may, es_conversion, num_fmt, color_hdr)
    # es_conversion=True => semaforo por ratio; False => semaforo por volumen
    indicadores = [
        # Volumenes
        ("RT  (Abr)",          rt_a,  None,  False, "#,##0",     C["rt_row"],  False),
        ("RT  (May)",          rt_m,  None,  False, "#,##0",     C["rt_row"],  False),
        ("RUS (Abr)",          rus_a, None,  False, "#,##0",     C["rus_row"], False),
        ("RUS (May)",          rus_m, None,  False, "#,##0",     C["rus_row"], False),
        ("ALTAS (Abr)",        alt_a, None,  False, "#,##0",     C["alt_row"], False),
        ("ALTAS (May)",        alt_m, None,  False, "#,##0",     C["alt_row"], False),
        # Separador visual
        ("", None, None, False, None, C["dlt"], True),
        # Conversiones con semaforo
        ("%Conv RT->RUS  Abr", rus_a, rt_a,  True,  '0.0"%"',   C["rus_row"], False),
        ("%Conv RT->RUS  May", rus_m, rt_m,  True,  '0.0"%"',   C["rus_row"], True),
        ("%Conv RUS->ALTA Abr",alt_a, rus_a, True,  '0.0"%"',   C["alt_row"], False),
        ("%Conv RUS->ALTA May",alt_m, rus_m, True,  '0.0"%"',   C["alt_row"], True),
        ("%Conv CON->ALTA Abr",alt_a, con_a, True,  '0.0"%"',   C["con_row"], False),
        ("%Conv CON->ALTA May",alt_m, con_m, True,  '0.0"%"',   C["con_row"], True),
    ]

    # Para las filas "May" necesitamos la serie "Abr" correspondiente para el semaforo
    abr_prev = {
        "%Conv RT->RUS  May":  (rus_a, rt_a),
        "%Conv RUS->ALTA May": (alt_a, rus_a),
        "%Conv CON->ALTA May": (alt_a, con_a),
        "RT  (May)":           (rt_a,  None),
        "RUS (May)":           (rus_a, None),
        "ALTAS (May)":         (alt_a, None),
    }

    cur = 3
    for (etq, s_num, s_den, es_conv, nfmt, bg, es_sep) in indicadores:
        if es_sep:
            for c in range(COL_LBL, COL_TOTAL + 1):
                cs(ws, cur, c, None, bg=C["dlt"], border=False)
            ws.row_dimensions[cur].height = 4
            cur += 1
            continue

        # Calcular valores por zonal
        vals = {}
        total_num = 0.0
        total_den = 0.0
        for z in zs_list:
            if es_conv:
                n = _safe(s_num.get(z) if s_num is not None else None)
                d = _safe(s_den.get(z) if s_den is not None else None)
                vals[z] = round(n / d * 100, 1) if (n and d) else None
                if n:
                    total_num += n
                if d:
                    total_den += d
            else:
                vals[z] = int(round(_safe(s_num.get(z)) or 0))
                total_num += vals[z]

        total_val = (
            round(total_num / total_den * 100, 1) if (es_conv and total_den) else
            int(total_num)
        )

        cs(ws, cur, COL_LBL, etq, bold=False, bg=bg, align="left", sz=9)

        for zi, z in enumerate(zs_list):
            v = vals[z]
            # Semaforo solo para filas "May"
            if etq in abr_prev:
                if es_conv:
                    s_a_num, s_a_den = abr_prev[etq]
                    n_a = _safe(s_a_num.get(z) if s_a_num is not None else None)
                    d_a = _safe(s_a_den.get(z) if s_a_den is not None else None)
                    v_a = round(n_a / d_a * 100, 1) if (n_a and d_a) else None
                    cell_bg = _sem_bg_conv(v, v_a)
                    cell_fc = _sem_fc_conv(v, v_a)
                else:
                    s_a_num, _ = abr_prev[etq]
                    v_a = _safe(s_a_num.get(z) if s_a_num is not None else None)
                    v_a = int(round(v_a)) if v_a is not None else None
                    cell_bg = _sem_bg_vol(v, v_a)
                    cell_fc = _sem_fc_conv(v, v_a)
            else:
                cell_bg, cell_fc = bg, "000000"

            cs(ws, cur, 2 + zi,
               v if v is not None else "-",
               bg=cell_bg, fc=cell_fc,
               num_fmt=nfmt if v is not None else None, sz=9)

        # Columna TOTAL
        if etq in abr_prev:
            if es_conv:
                s_a_num, s_a_den = abr_prev[etq]
                tn_a = sum(_safe(s_a_num.get(z)) or 0 for z in zs_list)
                td_a = sum(_safe(s_a_den.get(z)) or 0 for z in zs_list) if s_a_den is not None else 0
                tv_a = round(tn_a / td_a * 100, 1) if td_a else None
                t_bg = _sem_bg_conv(total_val, tv_a)
                t_fc = _sem_fc_conv(total_val, tv_a)
            else:
                s_a_num, _ = abr_prev[etq]
                tv_a = int(sum(_safe(s_a_num.get(z)) or 0 for z in zs_list))
                t_bg = _sem_bg_vol(total_val, tv_a)
                t_fc = _sem_fc_conv(total_val, tv_a)
        else:
            t_bg, t_fc = bg, "000000"

        cs(ws, cur, COL_TOTAL, total_val, bold=True, bg=t_bg, fc=t_fc,
           num_fmt=nfmt if total_val is not None else None, sz=9)

        ws.row_dimensions[cur].height = 14
        cur += 1

    # ── Leyenda ───────────────────────────────────────────────────────────────
    cur += 1
    cs(ws, cur, COL_LBL, "LEYENDA (filas Mayo):", bold=True, bg=None, border=False, sz=9)
    cur += 1
    for txt, bg in [
        ("VERDE  : Mayo >= Abril  (mantiene o mejora)", C["sem_verde"]),
        ("AMARILLO: Mayo entre 80-99% de Abril  (leve caida)", C["sem_amarillo"]),
        ("ROJO   : Mayo < 80% de Abril  (caida significativa)", C["sem_rojo"]),
    ]:
        cs(ws, cur, COL_LBL, txt, bold=False, bg=bg, align="left", border=False, sz=9)
        ws.row_dimensions[cur].height = 13
        cur += 1

    # ── Anchos ────────────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 26
    for zi in range(N):
        ws.column_dimensions[get_column_letter(2 + zi)].width = 9
    ws.column_dimensions[get_column_letter(COL_TOTAL)].width = 10
    ws.freeze_panes = "B3"


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    global DIA_CORTE, DIAS_EQUIV

    if not os.path.exists(BASE_DATOS_PATH):
        raise FileNotFoundError(
            f"No se encontro {BASE_DATOS_PATH}. "
            "Ejecuta primero: python generar_base_datos.py"
        )

    # Detectar el ultimo dia con datos en mayo y usarlo como corte
    DIA_CORTE  = _detectar_dia_corte()
    DIAS_EQUIV = _calcular_dias_equiv()
    print(f"Leyendo datos desde: {os.path.basename(BASE_DATOS_PATH)}  |  Dia corte: {DIA_CORTE}")

    datos   = _cargar_datos()
    zonales = _zonales_presentes(datos)

    wb = Workbook()
    wb.remove(wb.active)

    print("  Generando DATA...")
    ws_data = wb.create_sheet(title="DATA")
    _escribir_hoja_data(ws_data, datos, zonales)
    ws_data.sheet_state = "hidden"

    print("  Generando RESUMEN...")
    ws_res = wb.create_sheet(title="RESUMEN")
    _escribir_hoja_resumen(ws_res, datos, zonales)

    print("  Generando FUNNEL...")
    ws_fun = wb.create_sheet(title="FUNNEL")
    _escribir_hoja_funnel(ws_fun, datos, zonales)

    print("  Generando RT...")
    ws_rt = wb.create_sheet(title="RT")
    _escribir_hoja_metrica(ws_rt, "RT", datos, zonales)

    print("  Generando ALT...")
    ws_alt = wb.create_sheet(title="ALT")
    _escribir_hoja_metrica(ws_alt, "ALTAS", datos, zonales)

    print("  Generando RUS...")
    ws_rus = wb.create_sheet(title="RUS")
    _escribir_hoja_metrica(ws_rus, "RUS", datos, zonales)

    print("  Generando CON...")
    ws_con = wb.create_sheet(title="CON")
    _escribir_hoja_metrica(ws_con, "CON", datos, zonales)

    output_path = os.path.join(
        OUTPUT_DIR, f"DRILL_DOWN_ABR_vs_MAY_{DIA_CORTE:02d}.xlsx"
    )
    wb.save(output_path)
    print(f"\nArchivo generado: {output_path}")
    return output_path


if __name__ == "__main__":
    main()
