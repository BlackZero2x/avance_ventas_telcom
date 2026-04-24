"""
Helper para actualizar archivos Excel vía COM (win32com).

Fix 2026-03-25:
- refresh_and_save() ahora refresca TAMBIÉN las tablas dinámicas después
  de RefreshAll(). Antes solo se hacía en BacksProcess._process_fija(),
  quedando sin refresco en los archivos de Jesús y AVANCE_RESUMIDO.
- El orden es: deshabilitar BackgroundQuery → RefreshAll() → PivotCache.Refresh()
  → espera adaptativa → Save → Close.
"""
import logging
import os
import time

import win32com.client as win32
import pythoncom


class ExcelHelper:

    # Tiempo de espera adaptativo según el archivo
    _WAIT_EXTENDED_KEYWORDS = ["FIJA_RU_ALTAS", "MIGRAS"]  # 90s
    _WAIT_NORMAL = 30    # segundos para archivos estándar
    _WAIT_EXTENDED = 90  # segundos para archivos pesados con muchas conexiones

    def refresh_and_save(self, excel_path: str) -> bool:
        """
        Abre el archivo, deshabilita BackgroundQuery en todas las conexiones,
        ejecuta RefreshAll(), refresca todas las tablas dinámicas, espera,
        guarda y cierra.

        Retorna True si todo fue exitoso, False en caso de error.
        """
        excel = None
        try:
            pythoncom.CoInitialize()

            path = os.path.normpath(excel_path)
            filename = os.path.basename(path)

            if not os.path.exists(path):
                logging.error(f"Archivo no encontrado: {path}")
                return False

            logging.info(f"Actualizando Excel: {filename}")
            logging.info("Iniciando Excel en segundo plano...")

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            logging.info("Abriendo archivo...")
            wb = excel.Workbooks.Open(path)
            logging.info("Archivo abierto exitosamente")

            # ── 1. Deshabilitar BackgroundQuery ────────────────────
            for conn in wb.Connections:
                try:
                    if conn.Type == 1:   # xlConnectionTypeOLEDB
                        conn.OLEDBConnection.BackgroundQuery = False
                    elif conn.Type == 2:  # xlConnectionTypeODBC
                        conn.ODBCConnection.BackgroundQuery = False
                except Exception:
                    pass

            # ── 2. RefreshAll ──────────────────────────────────────
            logging.info("Refrescando conexiones de datos...")
            wb.RefreshAll()

            # ── 3. Refrescar tablas dinámicas ──────────────────────
            # RefreshAll() actualiza las fuentes de datos pero NO garantiza
            # que las tablas dinámicas recalculen antes de que guardemos.
            # PivotCache().Refresh() fuerza el recálculo hoja por hoja.
            pivot_count = 0
            for sheet in wb.Sheets:
                for pivot in sheet.PivotTables():
                    try:
                        pivot.PivotCache().Refresh()
                        pivot_count += 1
                    except Exception:
                        pass
            if pivot_count:
                logging.info(f"Tablas dinámicas refrescadas: {pivot_count}")

            # ── 4. Espera adaptativa ───────────────────────────────
            wait_seconds = self._WAIT_NORMAL
            for keyword in self._WAIT_EXTENDED_KEYWORDS:
                if keyword.upper() in filename.upper():
                    wait_seconds = self._WAIT_EXTENDED
                    break

            label = "extendido" if wait_seconds == self._WAIT_EXTENDED else "normal"
            logging.info(f"Archivo {label} detectado - esperando {wait_seconds}s para estabilizar...")

            elapsed = 0
            step = 30
            while elapsed < wait_seconds:
                remaining = min(step, wait_seconds - elapsed)
                time.sleep(remaining)
                elapsed += remaining
                logging.info(f"  Esperando... ({elapsed}s / {wait_seconds}s)")

            logging.info("Tiempo de espera completado")

            # ── 5. Guardar y cerrar ────────────────────────────────
            logging.info("Guardando archivo...")
            wb.Save()
            logging.info("Cerrando archivo...")
            wb.Close(SaveChanges=False)
            logging.info("Cerrando Excel...")
            excel.Quit()
            logging.info("Excel actualizado exitosamente")
            return True

        except Exception as e:
            logging.error(f"Error actualizando Excel: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False
        finally:
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass