#!/usr/bin/env python3
"""
Test del módulo cuadro_resumen_sup_process.py
Genera una imagen de cuadro para LETICIA/TRUJILLO y la envía a tu número.
"""
import sys
import os
from pathlib import Path

# Añadir módulos al path
sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent / "modules"))
sys.path.insert(0, str(Path(__file__).parent / "whatsapp_server"))

from modules.cuadro_resumen_sup_process import (
    _construir_tabla_zonal2,
    _capturar_imagen_tabla,
    _obtener_fecha_maxima
)
from wa_client import WhatsAppClient
import pandas as pd
from datetime import datetime

# Config
AVANCE_FILE = r"C:\proyectos\AVANCE_MOVISTAR\AVANCE_2026-07-26.xlsx"
TU_NUMERO = "51975155264@c.us"

def main():
    print("[TEST] Iniciando test de cuadro_resumen_sup_process.py")
    print(f"[TEST] Archivo: {AVANCE_FILE}")
    print(f"[TEST] Destino: {TU_NUMERO}")

    if not os.path.exists(AVANCE_FILE):
        print(f"[ERROR] Archivo no encontrado: {AVANCE_FILE}")
        return False

    try:
        # Cargar datos
        print("[TEST] Cargando datos de AVANCE...")
        df_rh = pd.read_excel(AVANCE_FILE, sheet_name='RH')
        df_rt = pd.read_excel(AVANCE_FILE, sheet_name='RT')
        df_altas = pd.read_excel(AVANCE_FILE, sheet_name='ALTAS')

        print(f"[TEST] RH: {len(df_rh)} filas")
        print(f"[TEST] RT: {len(df_rt)} filas")
        print(f"[TEST] ALTAS: {len(df_altas)} filas")

        # Obtener fecha máxima
        max_fecha = _obtener_fecha_maxima(df_rt)
        fecha_str = max_fecha.strftime("%d/%m/%Y") if max_fecha else datetime.now().strftime("%d/%m/%Y")
        print(f"[TEST] Fecha: {fecha_str}")

        # Construir cuadro para TRUJILLO
        print("[TEST] Construyendo cuadro para TRUJILLO...")
        tabla_trujillo = _construir_tabla_zonal2(df_rh, df_rt, df_altas, ['TRUJILLO'])

        if len(tabla_trujillo) == 0:
            print("[WARN] No hay datos para TRUJILLO")
            return False

        print(f"[TEST] Tabla TRUJILLO: {len(tabla_trujillo)} filas")
        print(tabla_trujillo)

        # Generar imagen
        print("[TEST] Generando imagen...")
        titulo = f"Cuadro TRUJILLO - Avance al {fecha_str}"
        png_path = _capturar_imagen_tabla(tabla_trujillo, titulo=titulo)

        if not png_path or not os.path.exists(png_path):
            print("[ERROR] Falló la generación de imagen")
            return False

        print(f"[OK] Imagen generada: {png_path}")

        # Enviar por WhatsApp
        print("[TEST] Conectando a WhatsApp...")
        wa = WhatsAppClient()

        mensaje = f"[TEST] Cuadro TRUJILLO - Avance al {fecha_str}\n(Tabla base: RRHH con conteos de RT/ALTAS)"
        print(f"[TEST] Enviando a {TU_NUMERO}...")
        wa.send_image(TU_NUMERO, png_path, caption=mensaje)

        print("[OK] Imagen enviada correctamente")

        # Limpiar
        try:
            os.remove(png_path)
        except:
            pass

        return True

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
