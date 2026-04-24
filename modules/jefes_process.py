"""
Proceso JEFES (Jefe de proyecto + Jefes de zona + Gerente comercial):
Fuente: archivo AVANCE_{fecha}.xlsx generado por AVANCE.py

1. Captura TDS!B4:V15  → imagen → grupo con mensaje + menciones
2. Captura TDS!Y4:AT16 → imagen → grupo con mensaje + menciones
3. Enviar SEGUIMIENTO_VDD_FIJA_dd-mm-aaaa.xlsx como adjunto con mensaje + menciones
"""
import logging
import os
import glob
import time
import subprocess
from datetime import datetime, timedelta

import xlwings as xw
import win32gui
from PIL import ImageGrab
from msg_utils import pick_variant


class JefesProcess:
    def __init__(self, config, wa):
        self.config = config
        self.wa = wa

    def execute(self):
        logging.info("=" * 60)
        logging.info("INICIANDO PROCESO JEFES")
        logging.info("=" * 60)

        archivo = self._buscar_avance()
        if not archivo:
            logging.error("No se encontro archivo AVANCE_*.xlsx")
            return False

        temp_dir = self.config["temp_dir"]
        os.makedirs(temp_dir, exist_ok=True)
        grupo = self.config["jefes_wa_group"]
        menciones = self.config.get("jefes_menciones") or None

        if not self._enviar_capturas(archivo, temp_dir, grupo):
            logging.warning("Hubo errores en las capturas TDS")

        if not self._enviar_seguimiento(grupo, menciones):
            logging.warning("No se pudo enviar el archivo SEGUIMIENTO_VDD")

        logging.info("PROCESO JEFES COMPLETADO")
        return True

    def _buscar_avance(self):
        directorio = self.config["archivos_avance_dir"]
        ayer = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        nombre_esperado = os.path.join(directorio, f"AVANCE_{ayer}.xlsx")
        if os.path.exists(nombre_esperado):
            return nombre_esperado
        logging.error(
            f"No se encontro AVANCE_{ayer}.xlsx en {directorio}. "
            f"Ejecuta AVANCE.py primero para generar el archivo del dia."
        )
        return None

    def _enviar_capturas(self, archivo, temp_dir, grupo):
        app = None
        wb = None
        try:
            # Cerrar instancias residuales de Excel
            subprocess.run(["taskkill", "/f", "/im", "EXCEL.EXE"], capture_output=True)
            time.sleep(2)

            logging.info(f"  Abriendo con xlwings: {os.path.basename(archivo)}")
            app = xw.App(visible=True, add_book=False)
            app.display_alerts = False
            wb = app.books.open(os.path.normpath(archivo), update_links=False, read_only=False)

            # Esperar a que Excel termine de calcular (max 120s)
            logging.info("  Esperando que Excel termine de calcular...")
            for i in range(120):
                try:
                    if app.api.CalculationState == 0:  # xlDone
                        logging.info(f"  Listo tras {i}s")
                        break
                except Exception:
                    pass
                time.sleep(1)

            app.api.Calculation = -4135  # xlCalculationManual
            time.sleep(2)

            excel_hwnd = self._get_excel_hwnd()
            sheet = wb.sheets["TDS"]
            sheet.activate()
            time.sleep(1)

            msg_captura1 = pick_variant(self.config.get("jefes_mensaje_captura1_variants"), self.config["jefes_mensaje_captura1"])
            msg_captura2 = pick_variant(self.config.get("jefes_mensaje_captura2_variants"), self.config["jefes_mensaje_captura2"])

            for rango, cap_path, caption in [
                (self.config["jefes_tds_rango1"], os.path.join(temp_dir, "captura_tds_1.png"), msg_captura1),
                (self.config["jefes_tds_rango2"], os.path.join(temp_dir, "captura_tds_2.png"), msg_captura2),
            ]:
                logging.info(f"  Capturando TDS!{rango}...")
                # Scroll al rango para que sea visible antes de CopyPicture
                sheet.range(rango).api.Select()
                app.api.ActiveWindow.ScrollIntoView(
                    sheet.range(rango).left,
                    sheet.range(rango).top,
                    sheet.range(rango).width,
                    sheet.range(rango).height,
                )
                time.sleep(1)
                sheet.range(rango).api.CopyPicture(Appearance=1, Format=2)
                time.sleep(3)

                capturado = False
                for intento in range(8):
                    try:
                        img = ImageGrab.grabclipboard()
                        if img:
                            img.save(cap_path, "PNG")
                            logging.info(f"  Guardada: {cap_path}")
                            self.wa.send_image(grupo, cap_path, caption=caption)
                            time.sleep(12)
                            capturado = True
                            break
                        else:
                            time.sleep(2)
                    except Exception as e:
                        logging.warning(f"  Intento {intento+1} grabclipboard: {e}")
                        time.sleep(2)

                if not capturado:
                    logging.error(f"  No se pudo capturar TDS!{rango}")

            return True

        except Exception as e:
            logging.error(f"Error en capturas: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False
        finally:
            if wb:
                try:
                    wb.close()
                except Exception:
                    pass
            if app:
                try:
                    app.quit()
                except Exception:
                    pass

    def _enviar_seguimiento(self, grupo, menciones):
        try:
            archivos_dir = self.config["archivos_avance_dir"]
            archivos = sorted(glob.glob(os.path.join(archivos_dir, "SEGUIMIENTO_VDD_FIJA_*.xlsx")), reverse=True)
            if not archivos:
                logging.error(f"No se encontro SEGUIMIENTO_VDD_FIJA en: {archivos_dir}")
                return False

            archivo = archivos[0]
            logging.info(f"  Enviando SEGUIMIENTO: {os.path.basename(archivo)}")

            msg_seg = pick_variant(self.config.get("jefes_mensaje_seguimiento_variants"), self.config["jefes_mensaje_seguimiento"])
            if menciones:
                self.wa.send_mention(grupo, msg_seg, menciones)
                time.sleep(10)
                self.wa.send_file(grupo, archivo, caption="")
            else:
                self.wa.send_file(grupo, archivo, caption=msg_seg)

            return True

        except Exception as e:
            logging.error(f"Error enviando SEGUIMIENTO: {e}")
            return False

    def _get_excel_hwnd(self):
        hwnd_found = [0]
        try:
            def _cb(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    t = win32gui.GetWindowText(hwnd)
                    if "Microsoft Excel" in t or ".xlsx" in t or ".xlsb" in t:
                        hwnd_found[0] = hwnd
                        return False
                return True
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass
        return hwnd_found[0]
