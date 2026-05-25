"""
Proceso ALERTA SUPERVISORES — Envío diario ~10 AM.

Lee VDD2 del último AVANCE_{fecha}.xlsx y por cada supervisor envía
a su contacto de WhatsApp (si está en config) un resumen de:
  - Vendedores INACTIVOS
  - Vendedores ACTIVOS SIN CIERRE
  - Top-3 con mayor RATIO_CON_A_ALT (señal predictiva de cierre)
  - Texto de acción sugerida

El resumen para JEFES incluye todas las zonales en una sola tabla.

Configuración requerida en config.json:
  "supervisor_wa_contacts": {
      "NOMBRE SUPERVISOR": "51XXXXXXXXX@c.us",
      ...
  }
  "supervisor_alert_jefes_group": "nombre o ID del grupo jefes"  (opcional)
"""
import logging
import os
import sys
from datetime import datetime, timedelta

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "whatsapp_server"))

from msg_utils import pick_variant


def _buscar_avance(directorio):
    ayer = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    candidatos = [
        os.path.join(directorio, f"AVANCE_{ayer}.xlsx"),
        os.path.join(directorio, f"AVANCE_{datetime.now().strftime('%Y-%m-%d')}.xlsx"),
    ]
    for p in candidatos:
        if os.path.exists(p):
            return p
    return None


def _leer_vdd2(archivo):
    try:
        df = pd.read_excel(archivo, sheet_name="VDD2", header=0)
        return _enriquecer_vdd2(df)
    except Exception as e:
        logging.error(f"[SupervisorAlert] No se pudo leer VDD2: {e}")
        return None


def _enriquecer_vdd2(df):
    """
    Calcula totales y estado del vendedor directamente desde las columnas de días,
    porque las columnas de resumen (TOTAL_RT, ALERTAS, ESTADO_VENDEDOR, etc.)
    son fórmulas Excel que pandas lee como cero/NaN.
    """
    import re as _re

    rt_cols  = sorted([c for c in df.columns if _re.match(r"RT_D\d+$",  c)])
    alt_cols = sorted([c for c in df.columns if _re.match(r"ALT_D\d+$", c)])
    con_cols = sorted([c for c in df.columns if _re.match(r"CON_D\d+$", c)])

    df["_TOTAL_RT"]  = df[rt_cols].fillna(0).sum(axis=1)  if rt_cols  else 0
    df["_TOTAL_ALT"] = df[alt_cols].fillna(0).sum(axis=1) if alt_cols else 0
    df["_TOTAL_CON"] = df[con_cols].fillna(0).sum(axis=1) if con_cols else 0

    rt_ult3  = rt_cols[-3:]  if len(rt_cols)  >= 3 else rt_cols
    alt_ult3 = alt_cols[-3:] if len(alt_cols) >= 3 else alt_cols
    con_ult3 = con_cols[-3:] if len(con_cols) >= 3 else con_cols

    df["_RT_ULT3"]  = df[rt_ult3].fillna(0).sum(axis=1)  if rt_ult3  else 0
    df["_ALT_ULT3"] = df[alt_ult3].fillna(0).sum(axis=1) if alt_ult3 else 0
    df["_CON_ULT3"] = df[con_ult3].fillna(0).sum(axis=1) if con_ult3 else 0

    def _estado(row):
        if row["_RT_ULT3"] == 0 and row["_ALT_ULT3"] == 0 and row["_CON_ULT3"] == 0:
            return "INACTIVO"
        if row["_ALT_ULT3"] == 0:
            return "ACTIVO SIN CIERRE"
        return "ACTIVO"

    df["_ESTADO"] = df.apply(_estado, axis=1)

    # Ratio CON/ALT de últimos 3 días, con cap en 1 (igual que fórmula Excel de VDD2)
    df["_RATIO"] = df.apply(
        lambda r: min(round(r["_CON_ULT3"] / r["_ALT_ULT3"], 2), 1) if r["_ALT_ULT3"] > 0 else None,
        axis=1,
    )

    return df


def _resumen_supervisor(df_sup, nombre_sup):
    """Genera texto de alerta para un supervisor dado su subconjunto de VDD2."""
    inactivos  = df_sup[df_sup["_ESTADO"] == "INACTIVO"]
    sin_cierre = df_sup[df_sup["_ESTADO"] == "ACTIVO SIN CIERRE"]

    total  = len(df_sup)
    n_inac = len(inactivos)
    n_sin  = len(sin_cierre)

    lines = [f"*Reporte diario — {datetime.now().strftime('%d/%m/%Y')}*"]
    lines.append(f"Supervisor: {nombre_sup}")
    lines.append(f"Vendedores activos: {total - n_inac - n_sin} / {total}")
    lines.append("")

    if n_inac:
        lines.append(f"*INACTIVOS ({n_inac}):*")
        for _, row in inactivos.iterrows():
            nombre = row.get("VENDEDOR", "?")
            antig  = row.get("ANTIG", "?")
            lines.append(f"  - {nombre} [{antig}]")
        lines.append("")

    if n_sin:
        lines.append(f"*ACTIVOS SIN CIERRE ({n_sin}):*")
        for _, row in sin_cierre.iterrows():
            nombre = row.get("VENDEDOR", "?")
            antig  = row.get("ANTIG", "?")
            ratio  = row.get("_RATIO", None)
            ratio_str = f" — ratio {ratio:.2f}" if isinstance(ratio, float) else ""
            lines.append(f"  - {nombre} [{antig}]{ratio_str}")
        lines.append("")

    # Top-3 por ratio CON→ALT (señal predictiva más fuerte)
    activos_con_ratio = df_sup[
        (df_sup["_ESTADO"] == "ACTIVO") &
        (df_sup["_RATIO"].notna())
    ].copy()
    if not activos_con_ratio.empty:
        top3 = activos_con_ratio.nlargest(3, "_RATIO")
        lines.append("*Top-3 ratio CON-ALT (mejor conversion):*")
        for _, row in top3.iterrows():
            lines.append(f"  - {row.get('VENDEDOR','?')}: {row['_RATIO']:.2f}")
        lines.append("")

    # Acción sugerida según distribución
    if n_inac > total * 0.3:
        lines.append("Accion: mas del 30% del equipo sin actividad. Revisar asistencia y pipeline.")
    elif n_sin > total * 0.4:
        lines.append("Accion: muchos vendedores registran pero no cierran. Acompanamiento en tecnica de cierre.")
    elif n_inac == 0 and n_sin == 0:
        lines.append("Equipo activo. Mantener ritmo.")

    return "\n".join(lines)


def _fecha_max_datos(df, archivo):
    """Devuelve la fecha del ultimo dia con datos en VDD2, en formato DD/MM/YYYY.
    La deriva del numero mas alto en las columnas RT_Dnn combinado con el
    mes/anio extraido del nombre del archivo (AVANCE_YYYY-MM-DD.xlsx).
    Si no puede determinarse, cae a ayer como fallback.
    """
    import re as _re
    try:
        rt_cols = [c for c in df.columns if _re.match(r"RT_D(\d+)$", c)]
        if rt_cols:
            ultimo_dia = max(int(_re.search(r"(\d+)$", c).group(1)) for c in rt_cols)
            # Extraer YYYY-MM del nombre del archivo
            m = _re.search(r"AVANCE_(\d{4})-(\d{2})-\d{2}\.xlsx$", os.path.basename(archivo))
            if m:
                anio, mes = int(m.group(1)), int(m.group(2))
                from datetime import date
                return date(anio, mes, ultimo_dia).strftime("%d/%m/%Y")
    except Exception:
        pass
    return (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")


def _resumen_ejecutivo_jefes(df, fecha_str):
    """Resumen ejecutivo de vendedores PLANILLA por zonal."""
    # Filtrar solo PLANILLA (PART-TIME ya fue reemplazado en AVANCE.py, pero por si acaso)
    df = df[df["ESQUEMA"].str.upper().str.strip().isin(["PLANILLA"])].copy()

    if df.empty:
        return None

    total_planilla = len(df)
    total_activos  = len(df[df["_ESTADO"].isin(["ACTIVO", "ACTIVO SIN CIERRE"])])
    total_sin_rt   = len(df[df["_ESTADO"] == "INACTIVO"])

    lines = [
        f"*Detalle del dia PLANILLA — {fecha_str}*",
        "",
        f"Total PLANILLA: {total_planilla} | Con venta (RT): {total_activos} | Sin venta: {total_sin_rt}",
        "",
        "Por ZONAL:",
    ]

    for zonal, grp in df.groupby("ZONAL"):
        total  = len(grp)
        con_rt = len(grp[grp["_ESTADO"].isin(["ACTIVO", "ACTIVO SIN CIERRE"])])
        sin_rt = len(grp[grp["_ESTADO"] == "INACTIVO"])

        ratios = grp["_RATIO"].dropna()
        ratio_prom = f"{ratios.mean():.2f}" if not ratios.empty else "-"

        if sin_rt == 0:
            etiqueta = "✅"
        elif sin_rt <= total * 0.2:
            etiqueta = "⚠️"
        else:
            etiqueta = "🛑"
        lines.append(
            f"  {etiqueta} {zonal}: {total} Vdds | {con_rt} con venta (RT) | {sin_rt} sin venta "
            f"| Ratio CON-->ALT: {ratio_prom}"
        )

    lines.append("")
    lines.append("Detalle del dia:")

    # Zonal con mas vendedores sin RT
    sin_rt_zonal = df[df["_ESTADO"] == "INACTIVO"].groupby("ZONAL").size()
    if not sin_rt_zonal.empty:
        peor = sin_rt_zonal.idxmax()
        n    = sin_rt_zonal[peor]
        lines.append(f"  - {peor}: {n} Vdds sin venta — accion urgente del sup")

    # Zonal con menor ratio de conversion
    activos_df = df[df["_ESTADO"] != "INACTIVO"].dropna(subset=["_RATIO"])
    ratio_por_zonal = activos_df.groupby("ZONAL")["_RATIO"].mean()
    if not ratio_por_zonal.empty:
        zonal_baja = ratio_por_zonal.idxmin()
        lines.append(f"  - {zonal_baja}: Ratio CON-->ALT mas bajo — posible problema de tecnica")

    return "\n".join(lines)


class SupervisorAlertProcess:
    def __init__(self, config, wa):
        self.config = config
        self.wa = wa

    def execute(self):
        logging.info("=" * 60)
        logging.info("INICIANDO PROCESO ALERTA SUPERVISORES")
        logging.info("=" * 60)

        archivo = _buscar_avance(self.config["archivos_avance_dir"])
        if not archivo:
            logging.error("[SupervisorAlert] No se encontro archivo AVANCE_.xlsx")
            return False

        df = _leer_vdd2(archivo)
        if df is None or df.empty:
            logging.error("[SupervisorAlert] VDD2 vacío o no legible")
            return False

        if "_ESTADO" not in df.columns:
            logging.error("[SupervisorAlert] _enriquecer_vdd2 no pudo calcular estados")
            return False

        contactos_sup = self.config.get("supervisor_wa_contacts", {})
        enviados = 0

        # Enviar alerta individual por supervisor
        for supervisor, wa_id in contactos_sup.items():
            df_sup = df[df["SUPERVISOR"].str.strip() == supervisor.strip()]
            if df_sup.empty:
                continue
            texto = _resumen_supervisor(df_sup, supervisor)
            try:
                self.wa.send_text(wa_id, texto)
                logging.info(f"[SupervisorAlert] Alerta enviada a {supervisor}")
                enviados += 1
            except Exception as e:
                logging.error(f"[SupervisorAlert] Error enviando a {supervisor}: {e}")

        # Enviar resumen ejecutivo al grupo de jefes
        grupo_jefes = self.config.get("supervisor_alert_jefes_group")
        if grupo_jefes:
            fecha_str = _fecha_max_datos(df, archivo)
            resumen = _resumen_ejecutivo_jefes(df, fecha_str)
            if resumen:
                try:
                    self.wa.send_text(grupo_jefes, resumen)
                    logging.info("[SupervisorAlert] Resumen ejecutivo enviado al grupo jefes")
                    enviados += 1
                except Exception as e:
                    logging.error(f"[SupervisorAlert] Error enviando resumen ejecutivo: {e}")

        logging.info(f"PROCESO ALERTA SUPERVISORES COMPLETADO — {enviados} mensajes enviados")
        return enviados > 0
