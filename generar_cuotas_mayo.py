"""
Genera cuotas_mayo2026.xlsx con cuotas por vendedor para mayo 2026.

Metodologia: Tipo 1 adaptado (Promedio/Maximo de 3 meses)
  RAW = ( PROMEDIO(M-3, M-2, M-1) + MAX(M-3, M-2, M-1) ) / 2

La distribucion individual respeta los totales zonales definidos en
cuotas_zonal_sup.xlsx hoja ZONAL (columnas CUOTA_MOVISTAR / CUOTA_MIFIBRA).
"""

from pathlib import Path
import numpy as np
import pandas as pd

BASE = Path(__file__).parent

# ── Rutas ───────────────────────────────────────────────────────────────────
VENDEDORES_FILE  = BASE / "VENDEDORES_ACTUALES.xlsx"
HISTORICO_FILE   = BASE / "altas_historico.xlsx"
ZONAL_FILE       = BASE / "cuotas_zonal_sup.xlsx"
MF_CSV           = Path(r"C:\Users\developer2\Documents\vpncompartido\BD_Ventas_AUREN.csv")
SALIDA           = BASE / "cuotas_mayo2026.xlsx"

CUOTA_MIN        = 10          # cuota minima por vendedor (donde la zona > 0)

# Periodos historicos: M-1=abr, M-2=mar, M-3=feb
PERIODOS_HIST = {
    "ALTAS_M1": (4, 2026),   # abr 2026
    "ALTAS_M2": (3, 2026),   # mar 2026
    "ALTAS_M3": (2, 2026),   # feb 2026
}
COL_HIST = {
    "ALTAS_M1": "altas_202604",
    "ALTAS_M2": "altas_202603",
    "ALTAS_M3": "altas_202602",
}


# ── Funciones auxiliares ─────────────────────────────────────────────────────

def distribuir_cuota(df_zona: pd.DataFrame, col_raw: str, cuota_total: int,
                     cuota_min: int = 0) -> pd.Series:
    """
    Distribuye `cuota_total` entre los vendedores de una zona usando `col_raw`
    como peso. Devuelve una Serie con la cuota entera para cada fila.

    - Si cuota_total == 0: todos reciben 0.
    - Redondeo largest-remainder para sumar exacto.
    - Aplica cuota_min: si alguno queda por debajo, se eleva al minimo y se
      descuenta el exceso del vendedor con mayor RAW.
    """
    n = len(df_zona)
    if n == 0 or cuota_total == 0:
        return pd.Series(0, index=df_zona.index)

    raw = df_zona[col_raw].fillna(0).clip(lower=0)

    # Si todos tienen RAW = 0, repartir igualitariamente
    if raw.sum() == 0:
        base = cuota_total // n
        resto = cuota_total % n
        cuotas = pd.Series(base, index=df_zona.index)
        # El resto se da al primer vendedor
        cuotas.iloc[:resto] += 1
        if cuota_min > 0:
            cuotas = cuotas.clip(lower=cuota_min)
        return cuotas.astype(int)

    peso = raw / raw.sum()
    exacto = peso * cuota_total

    # Metodo largest-remainder
    floor_v = exacto.astype(int)
    resto = cuota_total - floor_v.sum()
    fraccion = exacto - floor_v
    top_idx = fraccion.nlargest(int(resto)).index
    cuotas = floor_v.copy()
    cuotas.loc[top_idx] += 1

    # Aplicar minimo
    if cuota_min > 0:
        deficit = (cuotas < cuota_min).sum() * cuota_min - cuotas[cuotas < cuota_min].sum()
        cuotas[cuotas < cuota_min] = cuota_min
        # Descontar deficit del vendedor con mayor RAW (que no este en minimo)
        excedentes = cuotas[cuotas > cuota_min].sort_values(ascending=False)
        for idx in excedentes.index:
            quitar = min(deficit, cuotas[idx] - cuota_min)
            cuotas[idx] -= quitar
            deficit -= quitar
            if deficit <= 0:
                break

    return cuotas.astype(int)


def raw_tipo1(m3, m2, m1):
    """RAW = (promedio(M3,M2,M1) + max(M3,M2,M1)) / 2"""
    vals = [m3, m2, m1]
    return (np.mean(vals) + np.max(vals)) / 2


# ── 1. Cargar vendedores actuales ────────────────────────────────────────────
print("Cargando VENDEDORES_ACTUALES.xlsx...")
vdd = pd.read_excel(VENDEDORES_FILE)
vdd["DNI"] = pd.to_numeric(vdd["DNI"], errors="coerce").astype("Int64")
# Columna zona puede llamarse ZONA o ZONAL
if "ZONA" in vdd.columns and "ZONAL" not in vdd.columns:
    vdd = vdd.rename(columns={"ZONA": "ZONAL"})
print(f"  {len(vdd)} vendedores cargados, zonas: {sorted(vdd['ZONAL'].dropna().unique())}")


# ── 2. Cargar historico Movistar ─────────────────────────────────────────────
print("Cargando altas_historico.xlsx...")
hist = pd.read_excel(HISTORICO_FILE)
hist["dnivdd"] = pd.to_numeric(hist["dnivdd"], errors="coerce").astype("Int64")

# Verificar columnas disponibles
for alias, col in COL_HIST.items():
    if col not in hist.columns:
        print(f"  AVISO: columna '{col}' no encontrada en altas_historico — se usara 0")
        hist[col] = 0

hist_sel = hist[["dnivdd"] + list(COL_HIST.values())].rename(
    columns={**{"dnivdd": "DNI"}, **{v: k for k, v in COL_HIST.items()}}
)
print(f"  {len(hist_sel)} registros historicos")


# ── 3. Cargar historico MiFibra ──────────────────────────────────────────────
print("Cargando BD_Ventas_AUREN.csv (MiFibra)...")
try:
    mf_full = pd.read_csv(MF_CSV, encoding="utf-8-sig", low_memory=False)
    mf_full = mf_full[mf_full["ESTADO ORDEN SERVICIO 2"].str.strip() == "LIQUIDADA"].copy()
    mf_full["NUM DOC"] = pd.to_numeric(mf_full["NUM DOC"], errors="coerce").astype("Int64")

    mf_hist_rows = []
    for alias, (mes, anio) in PERIODOS_HIST.items():
        col_mf = alias.replace("ALTAS", "MF")   # MF_M1, MF_M2, MF_M3
        grp = (
            mf_full[(mf_full["MES_REG"] == mes) & (mf_full["AÑO_REG"] == anio)]
            .groupby("NUM DOC", as_index=False)
            .size()
            .rename(columns={"NUM DOC": "DNI", "size": col_mf})
        )
        mf_hist_rows.append(grp)

    mf_hist = mf_hist_rows[0]
    for df_tmp in mf_hist_rows[1:]:
        mf_hist = mf_hist.merge(df_tmp, on="DNI", how="outer")
    mf_hist = mf_hist.fillna(0)
    print(f"  {len(mf_hist)} vendedores con historial MiFibra")
except FileNotFoundError:
    print(f"  AVISO: CSV MiFibra no encontrado en {MF_CSV}. MiFibra se asignara 0.")
    mf_hist = pd.DataFrame(columns=["DNI", "MF_M1", "MF_M2", "MF_M3"])


# ── 4. Cargar objetivos zonales ──────────────────────────────────────────────
print("Cargando cuotas_zonal_sup.xlsx hoja ZONAL...")
zonal = pd.read_excel(ZONAL_FILE, sheet_name="ZONAL")
zonal.columns = [c.strip().upper() for c in zonal.columns]

# Filtrar por el periodo mas reciente disponible en la hoja ZONAL
if "PERIODO" in zonal.columns:
    periodo_zonal = zonal["PERIODO"].dropna().astype(str).max()
    zonal = zonal[zonal["PERIODO"].astype(str) == periodo_zonal].copy()
    print(f"  Periodo zonal: {periodo_zonal}")

print(f"  {len(zonal)} zonas cargadas")
print(zonal[["ZONAL", "CUOTA_MOVISTAR", "CUOTA_MIFIBRA"]].to_string(index=False))


# ── 5. Construir tabla maestra ───────────────────────────────────────────────
print("\nConstruyendo tabla maestra...")
df = (
    vdd[["DNI", "ZONAL", "VENDEDOR", "SUPERVISOR", "ESQUEMA"]]
    .merge(hist_sel, on="DNI", how="left")
    .merge(mf_hist,  on="DNI", how="left")
)

# Rellenar nulos con 0
for col in ["ALTAS_M1", "ALTAS_M2", "ALTAS_M3", "MF_M1", "MF_M2", "MF_M3"]:
    df[col] = df[col].fillna(0).astype(int)

# ── 6. Calcular RAW ─────────────────────────────────────────────────────────
df["PROMEDIO_3M"] = df[["ALTAS_M3", "ALTAS_M2", "ALTAS_M1"]].mean(axis=1).round(2)
df["MAXIMO_3M"]   = df[["ALTAS_M3", "ALTAS_M2", "ALTAS_M1"]].max(axis=1)
df["RAW_MOVISTAR"] = ((df["PROMEDIO_3M"] + df["MAXIMO_3M"]) / 2).round(4)

df["MF_PROMEDIO"] = df[["MF_M3", "MF_M2", "MF_M1"]].mean(axis=1).round(2)
df["MF_MAXIMO"]   = df[["MF_M3", "MF_M2", "MF_M1"]].max(axis=1)
df["RAW_MIFIBRA"]  = ((df["MF_PROMEDIO"] + df["MF_MAXIMO"]) / 2).round(4)

# Merge con objetivos zonales
df = df.merge(
    zonal[["ZONAL", "CUOTA_MOVISTAR", "CUOTA_MIFIBRA"]].rename(
        columns={"CUOTA_MOVISTAR": "CUOTA_ZONAL_MOV", "CUOTA_MIFIBRA": "CUOTA_ZONAL_MF"}
    ),
    on="ZONAL", how="left"
)
df["CUOTA_ZONAL_MOV"] = df["CUOTA_ZONAL_MOV"].fillna(0).astype(int)
df["CUOTA_ZONAL_MF"]  = df["CUOTA_ZONAL_MF"].fillna(0).astype(int)


# ── 7. Distribuir cuotas por zona ────────────────────────────────────────────
print("Distribuyendo cuotas por zona...")
df["CUOTA_MOVISTAR"] = 0
df["CUOTA_MIFIBRA"]  = 0

for zona, grp_idx in df.groupby("ZONAL").groups.items():
    grp = df.loc[grp_idx]
    cuota_mov = int(grp["CUOTA_ZONAL_MOV"].iloc[0])
    cuota_mf  = int(grp["CUOTA_ZONAL_MF"].iloc[0])

    # Movistar
    df.loc[grp_idx, "CUOTA_MOVISTAR"] = distribuir_cuota(
        grp, "RAW_MOVISTAR", cuota_mov, cuota_min=CUOTA_MIN if cuota_mov > 0 else 0
    ).values

    # MiFibra
    df.loc[grp_idx, "CUOTA_MIFIBRA"] = distribuir_cuota(
        grp, "RAW_MIFIBRA", cuota_mf, cuota_min=0
    ).values

# Peso zonal para mostrar en la hoja (calculado por zona)
df["PESO_ZONAL"] = 0.0
for zona, grp_idx in df.groupby("ZONAL").groups.items():
    raw_sum = df.loc[grp_idx, "RAW_MOVISTAR"].sum()
    if raw_sum > 0:
        df.loc[grp_idx, "PESO_ZONAL"] = (
            df.loc[grp_idx, "RAW_MOVISTAR"] / raw_sum
        ).round(4)


# ── 8. Construir resumen zonal ───────────────────────────────────────────────
resumen = (
    df.groupby("ZONAL")
    .agg(
        N_VENDEDORES        = ("DNI", "count"),
        SUMA_CUOTA_MOVISTAR = ("CUOTA_MOVISTAR", "sum"),
        OBJETIVO_MOVISTAR   = ("CUOTA_ZONAL_MOV", "first"),
        SUMA_CUOTA_MIFIBRA  = ("CUOTA_MIFIBRA", "sum"),
        OBJETIVO_MIFIBRA    = ("CUOTA_ZONAL_MF", "first"),
    )
    .reset_index()
)
resumen["DIF_MOVISTAR"] = resumen["SUMA_CUOTA_MOVISTAR"] - resumen["OBJETIVO_MOVISTAR"]
resumen["DIF_MIFIBRA"]  = resumen["SUMA_CUOTA_MIFIBRA"]  - resumen["OBJETIVO_MIFIBRA"]

# Fila total
total_row = pd.DataFrame([{
    "ZONAL":               "TOTAL",
    "N_VENDEDORES":        resumen["N_VENDEDORES"].sum(),
    "SUMA_CUOTA_MOVISTAR": resumen["SUMA_CUOTA_MOVISTAR"].sum(),
    "OBJETIVO_MOVISTAR":   resumen["OBJETIVO_MOVISTAR"].sum(),
    "SUMA_CUOTA_MIFIBRA":  resumen["SUMA_CUOTA_MIFIBRA"].sum(),
    "OBJETIVO_MIFIBRA":    resumen["OBJETIVO_MIFIBRA"].sum(),
    "DIF_MOVISTAR":        resumen["DIF_MOVISTAR"].sum(),
    "DIF_MIFIBRA":         resumen["DIF_MIFIBRA"].sum(),
}])
resumen = pd.concat([resumen, total_row], ignore_index=True)
print(resumen.to_string(index=False))


# ── 9. Construir hoja METODOLOGIA ────────────────────────────────────────────
metodologia_texto = [
    ["METODOLOGIA DE CALCULO DE CUOTAS — MAYO 2026"],
    [""],
    ["1. METODO BASE: Tipo 1 (Promedio/Maximo 3 meses)"],
    [""],
    ["   El metodo Tipo 1, aplicado originalmente en el proyecto SSFF, calcula una cuota"],
    ["   'raw' para cada vendedor que pondera su desempeno reciente:"],
    [""],
    ["   RAW = ( PROMEDIO(M-3, M-2, M-1) + MAX(M-3, M-2, M-1) ) / 2"],
    [""],
    ["   Donde:"],
    ["     M-1 = Altas de abril   2026  (mes de cierre mas reciente)"],
    ["     M-2 = Altas de marzo   2026"],
    ["     M-3 = Altas de febrero 2026"],
    [""],
    ["   Ventajas del metodo:"],
    ["   - Pondera el maximo historico reciente (aspiracional) con el promedio real."],
    ["   - Evita que un mes atipicamente alto infle toda la cuota."],
    ["   - Evita que un mes bajo por vacaciones o incidencias defina la cuota."],
    [""],
    ["2. DISTRIBUCION ZONAL"],
    [""],
    ["   Los objetivos zonales de mayo fueron fijados por la gerencia en"],
    ["   cuotas_zonal_sup.xlsx (hoja ZONAL). La distribucion individual"],
    ["   respeta estos totales exactamente:"],
    [""],
    ["   CUOTA_vdd = round( (RAW_vdd / SUM_RAW_zona) * CUOTA_TOTAL_zona )"],
    [""],
    ["   El redondeo usa el metodo 'largest remainder' para asegurar que la suma"],
    ["   de cuotas individuales sea identica al objetivo zonal."],
    [""],
    ["3. CUOTA MINIMA"],
    [""],
    ["   Todo vendedor en una zona con cuota > 0 recibe como minimo 10 altas."],
    ["   Si aplicar el minimo genera un excedente sobre el objetivo zonal, se"],
    ["   descuenta del vendedor con mayor RAW de esa zona."],
    [""],
    ["4. VENDEDORES SIN HISTORIAL"],
    [""],
    ["   Vendedores activos que no aparecen en altas_historico.xlsx reciben"],
    ["   RAW = 0. Al distribuir por zona, si todos los vendedores tienen RAW = 0,"],
    ["   la cuota se reparte igualitariamente. Luego se aplica el minimo de 10."],
    [""],
    ["5. MIFIBRA"],
    [""],
    ["   Para MiFibra se aplica la misma formula RAW usando las instalaciones"],
    ["   del CSV BD_Ventas_AUREN.csv (filtro: ESTADO ORDEN SERVICIO 2 = LIQUIDADA)."],
    ["   Solo 3 zonas tienen cuota MiFibra > 0: AREQUIPA (80), CHIMBOTE (20),"],
    ["   TRUJILLO (20). No se aplica cuota minima para MiFibra."],
    [""],
    ["6. FUENTES DE DATOS"],
    [""],
    ["   - Lista de vendedores : VENDEDORES_ACTUALES.xlsx"],
    ["   - Historico Movistar  : altas_historico.xlsx"],
    ["   - Historico MiFibra   : BD_Ventas_AUREN.csv"],
    ["   - Objetivos zonales   : cuotas_zonal_sup.xlsx (hoja ZONAL)"],
    [""],
    ["7. ACTUALIZACION MENSUAL"],
    [""],
    ["   Cada mes al cierre:"],
    ["   a) Agregar columna altas_YYYYMM a altas_historico.xlsx con las altas reales."],
    ["   b) Actualizar VENDEDORES_ACTUALES.xlsx con altas/bajas de vendedores."],
    ["   c) Actualizar cuotas_zonal_sup.xlsx con los nuevos objetivos zonales."],
    ["   d) Re-ejecutar este script para generar la nueva propuesta de cuotas."],
]


# ── 10. Escribir Excel ───────────────────────────────────────────────────────
print(f"\nEscribiendo {SALIDA}...")

# Columnas para hoja CUOTAS (orden presentacion)
cols_cuotas = [
    "ZONAL", "SUPERVISOR", "DNI", "VENDEDOR", "ESQUEMA",
    "ALTAS_M3", "ALTAS_M2", "ALTAS_M1",
    "PROMEDIO_3M", "MAXIMO_3M", "RAW_MOVISTAR", "PESO_ZONAL", "CUOTA_MOVISTAR",
    "MF_M3", "MF_M2", "MF_M1",
    "MF_PROMEDIO", "MF_MAXIMO", "RAW_MIFIBRA", "CUOTA_MIFIBRA",
]
df_out = df[cols_cuotas].sort_values(["ZONAL", "SUPERVISOR", "VENDEDOR"]).reset_index(drop=True)

with pd.ExcelWriter(SALIDA, engine="openpyxl") as writer:
    # Hoja CUOTAS
    df_out.to_excel(writer, sheet_name="CUOTAS", index=False)

    # Hoja RESUMEN_ZONAL
    resumen.to_excel(writer, sheet_name="RESUMEN_ZONAL", index=False)

    # Hoja METODOLOGIA
    met_df = pd.DataFrame(metodologia_texto, columns=["Descripcion"])
    met_df.to_excel(writer, sheet_name="METODOLOGIA", index=False, header=False)

    # ── Formato basico ───────────────────────────────────────────────────────
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    AZUL_HEADER = "1F4E79"
    GRIS_ALT    = "D9E1F2"
    VERDE_OK    = "C6EFCE"
    ROJO_MAL    = "FFC7CE"
    AMARILLO    = "FFEB9C"

    thin = Side(style="thin", color="BFBFBF")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    def estilo_header(ws, row=1, color=AZUL_HEADER):
        for cell in ws[row]:
            cell.font      = Font(bold=True, color="FFFFFF", size=10)
            cell.fill      = PatternFill("solid", fgColor=color)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = borde

    def autofit(ws, min_w=8, max_w=40):
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, min_w), max_w)

    # === Hoja CUOTAS ===
    ws_c = writer.sheets["CUOTAS"]
    ws_c.freeze_panes = "A2"
    ws_c.row_dimensions[1].height = 30
    estilo_header(ws_c)

    # Columnas de cuota resaltadas
    col_mov_idx = cols_cuotas.index("CUOTA_MOVISTAR") + 1
    col_mf_idx  = cols_cuotas.index("CUOTA_MIFIBRA") + 1

    for row_idx, row in enumerate(ws_c.iter_rows(min_row=2, max_row=ws_c.max_row), start=2):
        fill_base = PatternFill("solid", fgColor=GRIS_ALT) if row_idx % 2 == 0 else None
        for cell in row:
            if fill_base:
                cell.fill = fill_base
            cell.border = borde
            cell.alignment = Alignment(horizontal="center")

        # Resaltar cuota Movistar
        cell_mov = ws_c.cell(row=row_idx, column=col_mov_idx)
        cell_mov.font = Font(bold=True)
        cell_mov.fill = PatternFill("solid", fgColor="BDD7EE")

        # Resaltar cuota MiFibra
        cell_mf = ws_c.cell(row=row_idx, column=col_mf_idx)
        if cell_mf.value and cell_mf.value > 0:
            cell_mf.font = Font(bold=True)
            cell_mf.fill = PatternFill("solid", fgColor="E2EFDA")

    autofit(ws_c)

    # === Hoja RESUMEN_ZONAL ===
    ws_r = writer.sheets["RESUMEN_ZONAL"]
    ws_r.freeze_panes = "A2"
    ws_r.row_dimensions[1].height = 30
    estilo_header(ws_r, color="375623")

    max_r = ws_r.max_row
    dif_mov_col = list(resumen.columns).index("DIF_MOVISTAR") + 1
    dif_mf_col  = list(resumen.columns).index("DIF_MIFIBRA") + 1

    for row_idx, row in enumerate(ws_r.iter_rows(min_row=2, max_row=max_r), start=2):
        for cell in row:
            cell.border    = borde
            cell.alignment = Alignment(horizontal="center")

        is_total = ws_r.cell(row=row_idx, column=1).value == "TOTAL"
        if is_total:
            for cell in ws_r[row_idx]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor=AMARILLO)

        # Color en diferencia
        dif_mov = ws_r.cell(row=row_idx, column=dif_mov_col)
        dif_mf  = ws_r.cell(row=row_idx, column=dif_mf_col)
        for cell in [dif_mov, dif_mf]:
            if cell.value == 0:
                cell.fill = PatternFill("solid", fgColor=VERDE_OK)
            elif cell.value is not None and cell.value != 0:
                cell.fill = PatternFill("solid", fgColor=ROJO_MAL)

    autofit(ws_r)

    # === Hoja METODOLOGIA ===
    ws_m = writer.sheets["METODOLOGIA"]
    ws_m.column_dimensions["A"].width = 90
    ws_m.cell(1, 1).font = Font(bold=True, size=14, color=AZUL_HEADER)
    for row in ws_m.iter_rows(min_row=1, max_row=ws_m.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

print("[OK] Archivo generado:", SALIDA)

# Verificacion rapida
diffs = resumen[resumen["ZONAL"] != "TOTAL"][["DIF_MOVISTAR", "DIF_MIFIBRA"]]
if (diffs == 0).all().all():
    print("[OK] Todos los totales zonales cuadran exactamente.")
else:
    print("[ADVERTENCIA] Diferencias detectadas:")
    print(diffs)
