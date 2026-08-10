#!/usr/bin/env python3
"""
Test completo: Envía cuadros de TRUJILLO, CHIMBOTE y HUARAZ a tu número.
"""
import sys
import os
import time
from pathlib import Path

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

AVANCE_FILE = r"C:\proyectos\AVANCE_MOVISTAR\AVANCE_2026-07-26.xlsx"
TU_NUMERO = "51975155264@c.us"

def enviar_cuadro(wa, df_rh, df_rt, df_altas, region, destino, fecha_str):
    """Construye y envía un cuadro para una región."""
    print(f"\n[TEST] Procesando {region}...")

    try:
        # Construir tabla
        tabla = _construir_tabla_zonal2(df_rh, df_rt, df_altas, [region])

        if len(tabla) == 0:
            print(f"[WARN] No hay datos para {region}")
            return False

        print(f"[OK] Tabla {region}: {len(tabla)} vendedores")

        # Generar imagen
        titulo = f"Cuadro {region} - Avance al {fecha_str}"
        png_path = _capturar_imagen_tabla(tabla, titulo=titulo)

        if not png_path or not os.path.exists(png_path):
            print(f"[ERROR] Falló generación de imagen para {region}")
            return False

        # Enviar
        mensaje = f"[TEST] Cuadro {region} (RRHH como base, conteos de RT/ALTAS)"
        print(f"[OK] Enviando cuadro {region}...")
        wa.send_image(destino, png_path, caption=mensaje)
        print(f"[OK] {region} enviado")

        # Limpiar
        try:
            os.remove(png_path)
        except:
            pass

        return True

    except Exception as e:
        print(f"[ERROR] {region}: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("[TEST] Cuadro resumen - 3 regiones (TRUJILLO, CHIMBOTE, HUARAZ)")
    print(f"[TEST] Archivo: {AVANCE_FILE}")

    if not os.path.exists(AVANCE_FILE):
        print(f"[ERROR] Archivo no encontrado: {AVANCE_FILE}")
        return False

    try:
        # Cargar datos
        print("[TEST] Cargando datos...")
        df_rh = pd.read_excel(AVANCE_FILE, sheet_name='RH')
        df_rt = pd.read_excel(AVANCE_FILE, sheet_name='RT')
        df_altas = pd.read_excel(AVANCE_FILE, sheet_name='ALTAS')

        max_fecha = _obtener_fecha_maxima(df_rt)
        fecha_str = max_fecha.strftime("%d/%m/%Y") if max_fecha else datetime.now().strftime("%d/%m/%Y")

        print(f"[OK] Datos cargados (fecha: {fecha_str})")

        # Conectar WhatsApp
        wa = WhatsAppClient()

        # Enviar 3 cuadros
        regiones = ['TRUJILLO', 'CHIMBOTE', 'HUARAZ']
        resultados = {}

        for region in regiones:
            resultado = enviar_cuadro(wa, df_rh, df_rt, df_altas, region, TU_NUMERO, fecha_str)
            resultados[region] = resultado
            time.sleep(3)  # Pausa entre envíos

        # Resumen
        print("\n" + "="*50)
        print("[RESUMEN]")
        for region, ok in resultados.items():
            estado = "[OK]" if ok else "[FAIL]"
            print(f"  {region}: {estado}")

        total_ok = sum(1 for v in resultados.values() if v)
        print(f"\nResultado: {total_ok}/{len(regiones)} cuadros enviados")

        return total_ok == len(regiones)

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
