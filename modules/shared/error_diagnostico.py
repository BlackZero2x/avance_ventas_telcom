"""
Diagnóstico automático de errores de AVANCE.py.
Analiza el traceback y devuelve una sugerencia de solución legible para humanos.
"""
import os
import re


# Directorio raíz del proyecto
_PROYECTO_DIR = r"C:\proyectos\AVANCE_MOVISTAR"

# Cada regla: (patrón regex en el traceback, título, explicación, pasos de solución)
_REGLAS = [
    (
        r"UnicodeDecodeError|codec can'?t decode|'utf-?8'.*codec|charmap.*codec",
        "Error de codificación de caracteres (Unicode)",
        "Python no pudo leer un archivo porque tiene caracteres especiales (tildes, ñ, etc.) "
        "en una codificación distinta a la esperada. Ocurre frecuentemente con archivos Excel "
        "o CSV descargados de sistemas legados.",
        [
            "Identificar el archivo mencionado en el traceback (línea 'File ...').",
            "Si es un CSV: abrirlo en Excel → Guardar como → CSV UTF-8 (con BOM).",
            "Si es un xlsx: no debería tener este error — revisar si hay un pd.read_csv() "
            "sin encoding='utf-8-sig' en el script.",
            f"Archivo principal a revisar: {_PROYECTO_DIR}\\AVANCE.py",
            "Buscar en el código la línea indicada en el traceback y agregar encoding='utf-8-sig'.",
        ],
    ),
    (
        r"UnicodeEncodeError|'cp1252'|'ascii'.*codec.*encode",
        "Error de codificación al escribir (UnicodeEncodeError)",
        "Python intentó escribir texto con caracteres especiales (tildes, ñ) en un contexto "
        "que no soporta UTF-8, como la consola de Windows (cp1252) o un archivo abierto sin encoding.",
        [
            "Si el error es en la consola: es cosmético, no afecta el resultado. "
            "Agregar errors='replace' al print o logging.",
            "Si es al escribir un archivo: abrir con open(..., encoding='utf-8') explícito.",
            f"Archivo a revisar: {_PROYECTO_DIR}\\AVANCE.py (línea indicada en el traceback).",
        ],
    ),
    (
        r"FileNotFoundError|No such file or directory|cannot find the file",
        "Archivo no encontrado",
        "El script intentó abrir un archivo que no existe en la ruta esperada. "
        "Puede ser un Excel de entrada (rh.xlsx, lcf.xlsx, cuotas.xlsx) que no fue copiado al directorio.",
        [
            "Revisar el nombre del archivo mencionado en el traceback.",
            f"Verificar que el archivo existe en: {_PROYECTO_DIR}",
            "Los archivos de entrada requeridos son: rh.xlsx, lcf.xlsx (opcional: cuotas.xlsx, cuotas_zonal_sup.xlsx).",
            "Si falta un archivo, copiarlo al directorio del proyecto y volver a ejecutar.",
        ],
    ),
    (
        r"PermissionError|[Aa]ccess is denied|WinError 32|being used by another process",
        "Archivo bloqueado por otro proceso (Excel abierto)",
        "El script intentó escribir un archivo .xlsx que está abierto en Excel. "
        "Windows bloquea los archivos mientras están abiertos.",
        [
            "Cerrar Microsoft Excel completamente.",
            f"Verificar que ningún archivo AVANCE_*.xlsx esté abierto en {_PROYECTO_DIR}",
            "Volver a ejecutar main_v2.py o run_modulo.py.",
        ],
    ),
    (
        r"sqlalchemy|pyodbc|ODBC|SQL Server|OperationalError.*Login|Cannot open database"
        r"|TCP Provider|Named Pipes Provider|Login failed",
        "Error de conexión a SQL Server",
        "El script no pudo conectarse a la base de datos SQL Server. "
        "Puede ser que el servidor no esté disponible, las credenciales sean incorrectas, "
        "o el equipo no tenga acceso de red al servidor.",
        [
            "Verificar que el servidor SQL esté disponible: ping AUREN22.",
            "Verificar que ODBC Driver 17 for SQL Server esté instalado.",
            f"Revisar las credenciales en {_PROYECTO_DIR}\\.env: SQL_SERVER, SQL_USER, SQL_PASSWORD.",
            "Intentar conectar manualmente con SQL Server Management Studio para confirmar acceso.",
        ],
    ),
    (
        r"KeyError(?!.*config)|key.*not found|column.*not in index|not in dataframe",
        "Columna o clave no encontrada en los datos",
        "El script intentó acceder a una columna o clave que no existe en el DataFrame o diccionario. "
        "Puede deberse a un cambio en el formato de los datos de entrada (SQL, Excel o CSV).",
        [
            "Revisar el traceback para identificar qué columna falta (aparece entre comillas en KeyError: 'columna').",
            f"Abrir {_PROYECTO_DIR}\\AVANCE.py en la línea indicada.",
            "Verificar si los datos fuente (SQL o Excel) cambiaron de formato o nombre de columna.",
            "Si el campo tiene un nombre distinto en la fuente, actualizar el rename dict correspondiente.",
        ],
    ),
    (
        r"ModuleNotFoundError|ImportError|No module named",
        "Módulo Python no instalado",
        "Falta una librería Python requerida por el script.",
        [
            "Identificar el módulo faltante en el traceback (aparece como: No module named 'nombre').",
            r"Activar el entorno virtual: C:\proyectos\.venv\Scripts\activate",
            "Instalar el módulo faltante: pip install <nombre_modulo>",
            r"O instalar todos los requerimientos: pip install -r C:\proyectos\requirements.txt",
        ],
    ),
    (
        r"MemoryError|cannot allocate|out of memory",
        "Error de memoria insuficiente",
        "El proceso agotó la memoria RAM disponible, probablemente al cargar un DataFrame muy grande.",
        [
            "Cerrar otras aplicaciones pesadas (Excel, Chrome) y volver a intentar.",
            "Si el error persiste, puede ser que los datos SQL sean inusualmente grandes ese día.",
            f"Revisar en {_PROYECTO_DIR}\\AVANCE.py si hay carga de datos sin filtros.",
        ],
    ),
    (
        r"HttpError|403|401|invalid_grant|Token.*expired|credentials.*invalid"
        r"|google.*auth|oauth",
        "Error de autenticación con Google APIs",
        "El token de acceso a Google APIs expiró o es inválido. "
        "Requiere re-autenticación manual.",
        [
            f"Eliminar el archivo token.json en {_PROYECTO_DIR}",
            "Volver a ejecutar el script — abrirá un navegador para autenticarse.",
            "Confirmar con la cuenta augusto.moreno@auren.com.pe",
        ],
    ),
    (
        r"xlwings|Excel.*COM|win32com|pywintypes|COMError",
        "Error de automatización Excel (xlwings/COM)",
        "El script no pudo controlar Excel a través de la interfaz COM de Windows. "
        "Ocurre si Excel está cerrado cuando xlwings lo necesita, o si hay una instancia colgada.",
        [
            "Cerrar todas las ventanas de Excel.",
            "Abrir el Administrador de Tareas y terminar cualquier proceso EXCEL.EXE residual.",
            "Volver a ejecutar el script.",
            "Si persiste: reiniciar el equipo y ejecutar nuevamente.",
        ],
    ),
]

_REGLA_GENERICA = (
    "Error no reconocido automáticamente",
    "El script terminó con un error que no tiene una solución predefinida. "
    "Se adjunta el log completo para diagnóstico manual.",
    [
        f"Revisar el traceback completo en este correo.",
        f"Directorio del proyecto: {_PROYECTO_DIR}",
        f"Archivo principal: {_PROYECTO_DIR}\\AVANCE.py",
        "Identificar la línea exacta del error en el traceback (última entrada 'File ..., line N').",
    ],
)


def diagnosticar(stderr: str, stdout: str = "") -> dict:
    """
    Analiza el output de AVANCE.py y devuelve un dict con:
      - titulo: str
      - descripcion: str
      - pasos: list[str]
      - linea_error: str | None  (última línea de traceback tipo 'File X, line N')
      - tipo_error: str | None   (ej. 'UnicodeDecodeError: ...')
    """
    texto = (stderr or "") + "\n" + (stdout or "")

    linea_error = _extraer_linea_error(texto)
    tipo_error  = _extraer_tipo_error(texto)

    for patron, titulo, descripcion, pasos in _REGLAS:
        if re.search(patron, texto, re.IGNORECASE):
            return {
                "titulo":      titulo,
                "descripcion": descripcion,
                "pasos":       pasos,
                "linea_error": linea_error,
                "tipo_error":  tipo_error,
            }

    titulo, descripcion, pasos = _REGLA_GENERICA
    return {
        "titulo":      titulo,
        "descripcion": descripcion,
        "pasos":       pasos,
        "linea_error": linea_error,
        "tipo_error":  tipo_error,
    }


def _extraer_linea_error(texto: str) -> str | None:
    """Extrae la última línea 'File "...", line N, in ...' del traceback."""
    matches = re.findall(r'File "([^"]+)", line (\d+)', texto)
    if not matches:
        return None
    ruta, linea = matches[-1]
    nombre = os.path.basename(ruta)
    return f"{nombre}, línea {linea}  ({ruta})"


def _extraer_tipo_error(texto: str) -> str | None:
    """Extrae el nombre y mensaje del error (ej. 'UnicodeDecodeError: ...')."""
    # Buscar la última línea que empiece con un nombre de excepción Python
    matches = re.findall(
        r'^([A-Z][a-zA-Z]+(?:Error|Exception|Warning|Interrupt|Stop)[^\n]*)',
        texto, re.MULTILINE
    )
    if matches:
        return matches[-1].strip()[:300]
    return None


def formatear_html(diag: dict, stderr: str, log_path: str, proyecto_dir: str, hostname: str = "developer7") -> str:
    """Genera el cuerpo HTML del correo de diagnóstico."""

    pasos_html = "".join(f"<li>{p}</li>" for p in diag["pasos"])

    linea_bloque = ""
    if diag.get("linea_error"):
        linea_bloque = f"""
        <tr>
          <td style="padding:4px 8px;color:#555;white-space:nowrap;">Ubicación</td>
          <td style="padding:4px 8px;font-family:monospace;color:#c0392b;">{diag['linea_error']}</td>
        </tr>"""

    tipo_bloque = ""
    if diag.get("tipo_error"):
        tipo_bloque = f"""
        <tr>
          <td style="padding:4px 8px;color:#555;white-space:nowrap;">Tipo de error</td>
          <td style="padding:4px 8px;font-family:monospace;color:#c0392b;">{diag['tipo_error']}</td>
        </tr>"""

    # Últimas 60 líneas del stderr para no saturar el correo
    stderr_recortado = "\n".join((stderr or "").strip().splitlines()[-60:])

    return f"""
<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222;">

<h2 style="color:#c0392b;">&#9888; AVANCE.py terminó con error</h2>

<table style="border-collapse:collapse;margin-bottom:16px;">
  <tr>
    <td style="padding:4px 8px;color:#555;white-space:nowrap;">Problema detectado</td>
    <td style="padding:4px 8px;font-weight:bold;">{diag['titulo']}</td>
  </tr>{linea_bloque}{tipo_bloque}
  <tr>
    <td style="padding:4px 8px;color:#555;white-space:nowrap;">Equipo</td>
    <td style="padding:4px 8px;font-family:monospace;font-weight:bold;">{hostname}</td>
  </tr>
  <tr>
    <td style="padding:4px 8px;color:#555;white-space:nowrap;">Directorio del proyecto</td>
    <td style="padding:4px 8px;font-family:monospace;">{hostname}:{proyecto_dir}</td>
  </tr>
  <tr>
    <td style="padding:4px 8px;color:#555;white-space:nowrap;">Log completo</td>
    <td style="padding:4px 8px;font-family:monospace;">{hostname}:{log_path}</td>
  </tr>
</table>

<h3 style="color:#2c3e50;">Descripción</h3>
<p style="margin:0 0 12px;">{diag['descripcion']}</p>

<h3 style="color:#2c3e50;">Pasos para resolverlo</h3>
<ol style="margin:0 0 16px;padding-left:20px;">
{pasos_html}
</ol>

<h3 style="color:#2c3e50;">Detalle del error (últimas líneas)</h3>
<pre style="background:#f4f4f4;border:1px solid #ddd;padding:12px;font-size:12px;
            overflow-x:auto;white-space:pre-wrap;word-break:break-all;">{stderr_recortado}</pre>

<p style="color:#888;font-size:12px;margin-top:24px;">
  Mensaje generado automáticamente por el orquestador AVANCE MOVISTAR.<br>
  Los cambios deben realizarse en el equipo <strong>{hostname}</strong> &rarr; {proyecto_dir}
</p>
</body></html>
"""
