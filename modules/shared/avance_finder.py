"""
Función centralizada para localizar el archivo AVANCE_*.xlsx más reciente.

Lógica:
- Si config tiene 'periodo' (YYYY-MM), busca el AVANCE_*.xlsx más reciente
  cuya fecha esté dentro de ese mes, usando la fecha de modificación como desempate.
- Sin 'periodo', usa la lógica original: busca AVANCE_{ayer}.xlsx y
  AVANCE_{hoy}.xlsx como fallback (compatible con ejecución automática diaria).
"""
import glob
import logging
import os
from datetime import datetime, timedelta


def buscar_avance(directorio, config=None, prefijo="AVANCE_", sufijo=".xlsx"):
    """Retorna la ruta al archivo de avance más adecuado, o None si no se encuentra."""
    periodo = (config or {}).get("periodo", "")

    if periodo:
        # Periodo definido (YYYY-MM): tomar el más reciente del mes
        patron = os.path.join(directorio, f"{prefijo}????-??-??{sufijo}")
        candidatos = [
            p for p in glob.glob(patron)
            if _fecha_en_periodo(os.path.basename(p), prefijo, sufijo, periodo)
        ]
        if candidatos:
            # Más reciente por fecha en nombre (no por mtime, para ser determinista)
            archivo = max(candidatos, key=lambda p: _extraer_fecha(os.path.basename(p), prefijo, sufijo))
            logging.info(f"Archivo encontrado (periodo {periodo}): {archivo}")
            return archivo
        logging.error(
            f"No se encontro {prefijo}*.xlsx del periodo {periodo} en {directorio}. "
            f"Ejecuta AVANCE.py primero para generar el archivo."
        )
        return None

    # Sin periodo: comportamiento original (D-1, con fallback a hoy)
    ayer = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    hoy  = datetime.now().strftime("%Y-%m-%d")
    for fecha in (ayer, hoy):
        ruta = os.path.join(directorio, f"{prefijo}{fecha}{sufijo}")
        if os.path.exists(ruta):
            logging.info(f"Archivo encontrado: {ruta}")
            return ruta
    logging.error(
        f"No se encontro {prefijo}{ayer}.xlsx en {directorio}. "
        f"Ejecuta AVANCE.py primero para generar el archivo del dia."
    )
    return None


def _fecha_en_periodo(nombre, prefijo, sufijo, periodo):
    fecha_str = _extraer_fecha_str(nombre, prefijo, sufijo)
    return fecha_str.startswith(periodo)  # "2026-07-31" startswith "2026-07"


def _extraer_fecha(nombre, prefijo, sufijo):
    return _extraer_fecha_str(nombre, prefijo, sufijo)


def _extraer_fecha_str(nombre, prefijo, sufijo):
    # nombre = "AVANCE_2026-07-31.xlsx" → "2026-07-31"
    return nombre.replace(prefijo, "").replace(sufijo, "")
