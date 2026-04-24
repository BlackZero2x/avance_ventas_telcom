"""
Proceso BACKS:
1. Generar AVANCE_RESUMIDO.xlsx via generar_resumido.py (SQL directo, sin Excel COM)
2. Subir todas las hojas a Google Sheets (reemplaza el libro completo)
3. Notificar al grupo "Back de AUREN 2025" por WhatsApp (open-wa)
"""
import logging
import os
import subprocess
import sys
from datetime import datetime

import pandas as pd
from msg_utils import pick_variant


class BacksProcess:
    def __init__(self, config, sheets_service, wa):
        self.config = config
        self.sheets_service = sheets_service
        self.wa = wa

    def execute(self):
        logging.info("=" * 60)
        logging.info("INICIANDO PROCESO BACKS")
        logging.info("=" * 60)

        avance_path = self.config["avance_resumido_path"]
        periodo     = self.config.get("periodo", datetime.now().strftime("%Y-%m"))

        logging.info(f"[1/3] Generando AVANCE_RESUMIDO (periodo {periodo})...")
        script = os.path.join(os.path.dirname(os.path.dirname(__file__)), "generar_resumido.py")
        result = subprocess.run(
            [sys.executable, script, periodo],
            capture_output=True, text=True
        )
        if result.stdout:
            for line in result.stdout.strip().splitlines():
                logging.info(f"  {line}")
        if result.returncode != 0:
            if result.stderr:
                logging.error(result.stderr[-2000:])
            logging.error("Fallo generando AVANCE_RESUMIDO")
            return False

        logging.info("[2/3] Subiendo hojas a Google Sheets...")
        if not self._upload_to_sheets(avance_path):
            logging.error("Fallo subiendo a Google Sheets")
            return False

        logging.info("[3/3] Notificando grupo BACKS por WhatsApp...")
        mensaje = pick_variant(self.config.get("backs_message_variants"), self.config["backs_message"])
        self.wa.send_text(self.config["backs_wa_group"], mensaje)

        logging.info("PROCESO BACKS COMPLETADO")
        return True

    def _upload_to_sheets(self, excel_path):
        try:
            sheet_id = self.config["backs_google_sheet_id"]
            excel_file = pd.ExcelFile(excel_path)
            excel_sheet_names = excel_file.sheet_names
            logging.info(f"  {len(excel_sheet_names)} hojas encontradas en AVANCE_RESUMIDO")

            spreadsheet = self.sheets_service.spreadsheets().get(
                spreadsheetId=sheet_id
            ).execute()

            existing_sheets = {
                s["properties"]["title"]: s["properties"]["sheetId"]
                for s in spreadsheet.get("sheets", [])
            }

            requests = []
            for title, sid in existing_sheets.items():
                if title not in excel_sheet_names:
                    requests.append({"deleteSheet": {"sheetId": sid}})
            for name in excel_sheet_names:
                if name not in existing_sheets:
                    requests.append({"addSheet": {"properties": {"title": name}}})

            if requests:
                self.sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=sheet_id, body={"requests": requests}
                ).execute()

            for sheet_name in excel_sheet_names:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
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
                    range=f"{sheet_name}!A1:ZZ100000"
                ).execute()

                self.sheets_service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range=f"{sheet_name}!A1",
                    valueInputOption="USER_ENTERED",
                    body={"values": values}
                ).execute()
                logging.info(f"  Hoja '{sheet_name}' subida ({len(df)} filas)")

            logging.info("  Todas las hojas subidas exitosamente")
            return True

        except Exception as e:
            logging.error(f"Error subiendo a Google Sheets: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False
