"""
Proceso JEFES (Jefe de proyecto + Jefes de zona + Gerente comercial):
Fuente: archivo AVANCE_{fecha}.xlsx generado por AVANCE.py

1. Captura TDS!B4:V15  → imagen → grupo con mensaje + menciones
2. Captura TDS!Y4:AT17 → imagen → grupo con mensaje + menciones
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

    def _capturar_rango(self, app, wb, sheet, rango, cap_path, excel_hwnd):
        """Exporta un rango de Excel como PNG. Intenta CopyPicture+clipboard primero;
        si falla (sesiones sin escritorio interactivo), usa el método Chart como fallback."""

        # — Intento 1: CopyPicture → clipboard —
        for intento in range(3):
            try:
                excel_hwnd = self._get_excel_hwnd() or excel_hwnd
                if excel_hwnd:
                    try:
                        win32gui.ShowWindow(excel_hwnd, 9)
                        win32gui.SetForegroundWindow(excel_hwnd)
                    except Exception:
                        pass
                sheet.range(rango).api.Select()
                app.api.ActiveWindow.ScrollIntoView(
                    sheet.range(rango).left,
                    sheet.range(rango).top,
                    sheet.range(rango).width,
                    sheet.range(rango).height,
                )
                time.sleep(1 + intento)
                sheet.range(rango).api.CopyPicture(Appearance=1, Format=2)
                time.sleep(2)
                img = ImageGrab.grabclipboard()
                if img:
                    img.save(cap_path, "PNG")
                    logging.info(f"  Guardada (CopyPicture): {cap_path}")
                    return True
            except Exception as e:
                logging.warning(f"  CopyPicture intento {intento+1} fallido: {e}")
                time.sleep(2)

        # — Intento 2: Chart export (funciona sin escritorio interactivo) —
        logging.info(f"  CopyPicture agotado — usando Chart export para {rango}...")
        try:
            xl_range = sheet.range(rango)
            xl_range.api.Copy()
            time.sleep(1)

            # Crear un chart temporal en la misma hoja y pegarle la imagen del rango
            charts = sheet.api.ChartObjects()
            chart_obj = charts.Add(0, 0, xl_range.width, xl_range.height)
            chart = chart_obj.Chart
            chart.Paste()
            time.sleep(1)

            # Exportar el chart como PNG
            chart.Export(os.path.normpath(cap_path))
            chart_obj.Delete()
            time.sleep(0.5)

            if os.path.exists(cap_path):
                logging.info(f"  Guardada (Chart export): {cap_path}")
                return True
        except Exception as e:
            logging.error(f"  Chart export fallido para {rango}: {e}")

        return False

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

            # Traer Excel al frente — CopyPicture falla si Excel no es la ventana activa.
            # Usamos tanto la API COM de Excel como win32gui para maximizar las chances
            # en ejecuciones desatendidas (Programador de Tareas).
            try:
                app.api.Visible = True
                app.api.ActiveWindow.Activate()
            except Exception:
                pass
            if excel_hwnd:
                try:
                    import ctypes
                    # AllowSetForegroundWindow permite traer al frente desde procesos sin foco
                    ctypes.windll.user32.AllowSetForegroundWindow(ctypes.windll.kernel32.GetCurrentProcessId())
                    win32gui.ShowWindow(excel_hwnd, 9)   # SW_RESTORE
                    win32gui.SetForegroundWindow(excel_hwnd)
                    time.sleep(1)
                except Exception as e:
                    logging.warning(f"  No se pudo traer Excel al frente: {e}")

            msg_captura1 = pick_variant(self.config.get("jefes_mensaje_captura1_variants"), self.config["jefes_mensaje_captura1"])
            msg_captura2 = pick_variant(self.config.get("jefes_mensaje_captura2_variants"), self.config["jefes_mensaje_captura2"])
            menciones = self.config.get("jefes_menciones") or []

            for rango, cap_path, msg in [
                (self.config["jefes_tds_rango1"], os.path.join(temp_dir, "captura_tds_1.png"), msg_captura1),
                (self.config["jefes_tds_rango2"], os.path.join(temp_dir, "captura_tds_2.png"), msg_captura2),
            ]:
                logging.info(f"  Capturando TDS!{rango}...")
                if self._capturar_rango(app, wb, sheet, rango, cap_path, excel_hwnd):
                    # Enviar imagen primero (sin caption para no mezclar texto con menciones)
                    self.wa.send_image(grupo, cap_path, caption="")
                    time.sleep(5)
                    # Enviar texto como mención real para que WhatsApp lo renderice correctamente
                    if menciones:
                        self.wa.send_mention(grupo, msg, menciones)
                    else:
                        self.wa.send_text(grupo, msg)
                    time.sleep(12)
                else:
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
            candidatos = glob.glob(os.path.join(archivos_dir, "SEGUIMIENTO_VDD_FIJA_*.xlsx"))
            if not candidatos:
                logging.error(f"No se encontro SEGUIMIENTO_VDD_FIJA en: {archivos_dir}")
                return False

            def _fecha_desde_nombre(path):
                nombre = os.path.basename(path)
                # Formato: SEGUIMIENTO_VDD_FIJA_dd-mm-yyyy.xlsx
                partes = nombre.replace("SEGUIMIENTO_VDD_FIJA_", "").replace(".xlsx", "").split("-")
                try:
                    return datetime(int(partes[2]), int(partes[1]), int(partes[0]))
                except Exception:
                    return datetime.min

            archivo = max(candidatos, key=_fecha_desde_nombre)
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
