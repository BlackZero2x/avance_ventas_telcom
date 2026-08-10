"""
Proceso ITALO:
1. Buscar AVANCE_VENTAS_APPVENTORY_{fecha_ayer}.xlsx en Archivos_Avance
2. Leer hojas MES y DIA del libro Excel
3. Subir cada hoja al Google Sheet de Italo (limpiar y reemplazar)
4. Notificar a Italo por WhatsApp
"""
import logging
import os
import glob
from datetime import datetime, timedelta

import pandas as pd
from dotenv import load_dotenv
from msg_utils import pick_variant
from shared.avance_finder import buscar_avance

load_dotenv()


class ItaloProcess:
    def __init__(self, config, sheets_service, wa):
        self.config = config
        self.sheets_service = sheets_service
        self.wa = wa

    def execute(self):
        logging.info("=" * 50)
        logging.info("INICIANDO PROCESO ITALO")
        logging.info("=" * 50)

        logging.info("[1/3] Buscando archivo AVANCE_VENTAS_APPVENTORY...")
        archivo = self._buscar_appventory()
        if not archivo:
            logging.error("No se encontro AVANCE_VTAS_APPVENTORY_*.xlsx")
            return False

        logging.info(f"  Archivo encontrado: {os.path.basename(archivo)}")

        logging.info("[2/2] Subiendo hojas MES y DIA a Google Sheets...")
        sheet_id = os.environ.get("ITALO_SHEET_ID") or self.config.get("italo_google_sheet_id")
        if not sheet_id:
            logging.error("ITALO_SHEET_ID no definido en .env")
            return False
        ok = True
        for hoja in ["MES", "DIA"]:
            if not self._upload_sheet(archivo, hoja, sheet_id):
                ok = False

        if not ok:
            logging.error("Fallo al subir una o mas hojas a Google Sheets")
            return False

        logging.info("PROCESO ITALO COMPLETADO (sin notificacion WA)")
        return True

    def _buscar_appventory(self):
        return buscar_avance(
            self.config["archivos_avance_dir"],
            self.config,
            prefijo="AVANCE_VTAS_APPVENTORY_",
        )

    def _upload_sheet(self, archivo, hoja, sheet_id):
        try:
            df = pd.read_excel(archivo, sheet_name=hoja)
            logging.info(f"  Hoja '{hoja}': {len(df)} filas x {len(df.columns)} columnas")

            rows = []
            for _, row in df.iterrows():
                row_values = []
                for v in row:
                    if pd.isna(v):
                        row_values.append("")
                    elif isinstance(v, (pd.Timestamp, datetime)):
                        row_values.append(v.strftime("%Y-%m-%d"))
                    elif isinstance(v, (int, float)):
                        row_values.append(v)
                    else:
                        row_values.append(str(v))
                rows.append(row_values)

            values = [df.columns.tolist()] + rows

            self.sheets_service.spreadsheets().values().clear(
                spreadsheetId=sheet_id,
                range=f"{hoja}!A1:ZZ1000000"
            ).execute()

            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=f"{hoja}!A1",
                valueInputOption="USER_ENTERED",
                body={"values": values}
            ).execute()

            logging.info(f"  [OK] Hoja '{hoja}' actualizada en Google Sheets")
            return True

        except Exception as e:
            logging.error(f"  Error subiendo hoja '{hoja}': {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False
