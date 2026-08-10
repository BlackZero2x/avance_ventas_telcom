"""
Cargador BBDD Fija -> Google Sheet
Herramienta standalone para Katerine Palacios.

Flujo:
1. Ella hace doble clic en Cargador_BBDD_Fija.exe
2. Se abre un dialogo para elegir su archivo "BDD...xlsx" (parte en la carpeta
   de red predeterminada)
3. Se lee la hoja "BBDD Fija" completa
4. Se sube como VALORES (sin formulas ni formato) al Google Sheet, reemplazando
   todo el contenido de la pestana destino (GID fijo)
5. Pop-up de confirmacion o de error
"""

import os
import sys
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox

import gspread
from google.oauth2.service_account import Credentials

# ── Configuracion ────────────────────────────────────────────────────────
NOMBRE_USUARIO = "Katerine Palacios"

CARPETA_PREDETERMINADA = (
    r"\\SERVER\compartido\Reclutamiento\Reclutamiento General"
    r"\PROYECTOS\PROYECTO FIJA\Base Fija\FIJA"
)
PATRON_ARCHIVO = "BDD*.xlsx"
HOJA_ORIGEN = "BBDD Fija"

SHEET_ID = "1UwIkcOq_pqDRHfMrsH-RnGWRQqdgUpNLytsaX_RHmyc"
GID_DESTINO = 407197609

# Nombre del JSON de la cuenta de servicio. Debe estar en la misma carpeta
# que el .exe (PyInstaller lo busca junto al ejecutable, ver resource_path).
CREDENCIAL_SERVICIO = "service_account.json"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def ruta_junto_al_exe(nombre_archivo: str) -> str:
    """Resuelve una ruta relativa a la carpeta del .exe (o del script, en dev)."""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, nombre_archivo)


def elegir_archivo() -> str | None:
    carpeta_inicial = (
        CARPETA_PREDETERMINADA if os.path.isdir(CARPETA_PREDETERMINADA) else os.path.expanduser("~")
    )
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo BDD a cargar",
        initialdir=carpeta_inicial,
        filetypes=[("Archivos BDD Fija", PATRON_ARCHIVO), ("Excel", "*.xlsx")],
    )
    return ruta or None


def leer_hoja_como_valores(ruta_excel: str):
    """Lee la hoja HOJA_ORIGEN con openpyxl y devuelve una lista de listas
    (solo valores, celdas vacias como cadena vacia)."""
    import openpyxl

    wb = openpyxl.load_workbook(ruta_excel, data_only=True, read_only=True)
    if HOJA_ORIGEN not in wb.sheetnames:
        disponibles = ", ".join(wb.sheetnames)
        raise ValueError(
            f"No se encontro la hoja '{HOJA_ORIGEN}' en el archivo.\n"
            f"Hojas disponibles: {disponibles}"
        )
    ws = wb[HOJA_ORIGEN]

    filas = []
    for fila in ws.iter_rows(values_only=True):
        if all(celda is None for celda in fila):
            continue
        filas.append(["" if celda is None else celda for celda in fila])
    wb.close()

    if not filas:
        raise ValueError(f"La hoja '{HOJA_ORIGEN}' esta vacia.")
    return filas


def subir_a_sheet(filas):
    ruta_cred = ruta_junto_al_exe(CREDENCIAL_SERVICIO)
    if not os.path.isfile(ruta_cred):
        raise FileNotFoundError(
            f"No se encontro el archivo de credenciales '{CREDENCIAL_SERVICIO}' "
            f"junto al programa.\nRuta esperada: {ruta_cred}"
        )

    creds = Credentials.from_service_account_file(ruta_cred, scopes=SCOPES)
    cliente = gspread.authorize(creds)

    libro = cliente.open_by_key(SHEET_ID)
    hoja_destino = None
    for hoja in libro.worksheets():
        if hoja.id == GID_DESTINO:
            hoja_destino = hoja
            break
    if hoja_destino is None:
        raise ValueError(f"No se encontro ninguna pestana con GID {GID_DESTINO} en el Sheet destino.")

    hoja_destino.clear()

    num_filas = len(filas)
    num_cols = max(len(f) for f in filas)
    rango = gspread.utils.rowcol_to_a1(num_filas, num_cols)
    hoja_destino.update(
        range_name=f"A1:{rango}",
        values=filas,
        value_input_option="RAW",
    )
    return num_filas, num_cols


def main():
    root = tk.Tk()
    root.withdraw()

    ruta_excel = elegir_archivo()
    if not ruta_excel:
        return

    try:
        filas = leer_hoja_como_valores(ruta_excel)
        num_filas, num_cols = subir_a_sheet(filas)

        ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
        messagebox.showinfo(
            "Carga completada",
            f"{NOMBRE_USUARIO}: carga completada correctamente.\n\n"
            f"Archivo: {os.path.basename(ruta_excel)}\n"
            f"Filas cargadas: {num_filas}\n"
            f"Fecha/hora: {ahora}",
        )
    except Exception as exc:
        detalle = traceback.format_exc()
        try:
            with open(ruta_junto_al_exe("error_ultima_carga.log"), "w", encoding="utf-8") as f:
                f.write(detalle)
        except OSError:
            pass
        messagebox.showerror(
            "Error al cargar",
            f"{NOMBRE_USUARIO}, ocurrio un error y no se pudo completar la carga.\n\n"
            f"{exc}\n\n"
            f"(Detalle guardado en error_ultima_carga.log)",
        )


if __name__ == "__main__":
    main()
