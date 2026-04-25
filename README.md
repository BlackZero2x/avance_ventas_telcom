# Reporte y Sistema de Seguimiento de Ventas

Sistema de automatización end-to-end para el seguimiento diario de ventas de la línea fija de Movistar Perú. Consolida datos de múltiples fuentes, genera dashboards Excel con indicadores de gestión y los distribuye automáticamente a los integrantes del equipo comercial por WhatsApp y Google Sheets, sin intervención manual.

---

## Tabla de contenidos

1. [Para qué sirve este sistema](#para-qué-sirve-este-sistema)
2. [Cómo funciona en el día a día](#cómo-funciona-en-el-día-a-día)
3. [Arquitectura general](#arquitectura-general)
4. [Módulos y responsabilidades](#módulos-y-responsabilidades)
5. [Estructura del proyecto](#estructura-del-proyecto)
6. [Instalación desde cero](#instalación-desde-cero)
7. [Configuración](#configuración)
8. [Ejecución manual](#ejecución-manual)
9. [Secciones del ETL (AVANCE.py)](#secciones-del-etl-avancepy)
10. [Salidas que genera el sistema](#salidas-que-genera-el-sistema)
11. [Servidor y cliente de WhatsApp](#servidor-y-cliente-de-whatsapp)
12. [Mitigación de riesgo de ban en WhatsApp](#mitigación-de-riesgo-de-ban-en-whatsapp)
13. [Stack tecnológico](#stack-tecnológico)
14. [Glosario del dominio](#glosario-del-dominio)
15. [Preguntas frecuentes y solución de problemas](#preguntas-frecuentes-y-solución-de-problemas)

---

## Para qué sirve este sistema

Cada día hábil, el equipo de ventas de línea fija necesita saber:

- **¿Cuántas altas se registraron?** (clientes activados ese día)
- **¿Cómo va el avance vs la cuota?** (porcentaje de cumplimiento por supervisor/zona)
- **¿Qué vendedores están en riesgo?** (rendimiento bajo con semáforo rojo/amarillo/verde)
- **¿Cuál es el detalle individual por vendedor?** (VDD: ventas, antigüedad, score de riesgo)

Este sistema genera esa información de forma automática y la entrega a las personas correctas (jefes de zona, gerentes, analistas) a través de WhatsApp y Google Sheets, **sin que nadie tenga que ejecutar nada manualmente**.

---

## Cómo funciona en el día a día

El proceso completo se dispara con un único email y ocurre sin intervención:

```
<<<<<<< HEAD
SQL Server (eAuren) + Excel (rh, lcf, cuotas)
        │
        ▼
    AVANCE.py  ── ETL de ~3.500 líneas, 9 secciones
        │
        ▼
    Reportes Excel
    ├── AVANCE_{fecha}.xlsx         (dashboard principal)
    ├── AVANCE_RESUMIDO.xlsx        (resumen para Google Sheets)
    └── SEGUIMIENTO_VDD_FIJA.xlsx   (snapshot diario por vendedor)
        │
        ▼
    main_v2.py  ── Orquestador (Programador de Tareas Windows)
        │
        ├── BacksProcess    → Google Sheets + grupo WhatsApp
        ├── JefesProcess    → capturas TDS + archivo → grupo WhatsApp
        ├── JesusProcess    → archivo → contacto WhatsApp
        ├── CristianProcess → archivo → contacto WhatsApp
        ├── GuillermoProcess → archivo → contacto WhatsApp
        ├── CarlosProcess   → archivo → contacto WhatsApp
        └── ItaloProcess    → Google Sheets + contacto WhatsApp
                │
                ▼
        wa_client.py → wa_server.js (puerto 8002) → WhatsApp
=======
08:30 AM  →  El Programador de Tareas de Windows inicia main_v2.py
             y verifica que el servidor de WhatsApp esté activo.

08:30–??  →  main_v2.py revisa el correo cada 10 minutos,
             buscando el email que indica que los datos del día
             ya están disponibles en el sistema.

Al llegar →  AVANCE.py se ejecuta:
el email       - Consulta SQL Server para obtener los registros de altas
               - Cruza con archivos Excel de RH, riesgo crediticio y cuotas
               - Calcula KPIs, semáforos y tablas dinámicas
               - Genera los archivos Excel finales

Después  →   7 procesos de distribución corren en cadena:
               1. BACKS   → sube resumen a Google Sheets + notifica al equipo
               2. JEFES   → captura imágenes del dashboard TDS + envía a grupo
               3. JESUS   → envía AVANCE_{fecha}.xlsx por WhatsApp
               4. CRISTIAN→ envía AVANCE_{fecha}.xlsx por WhatsApp
               5. GUILLERMO→ envía AVANCE_{fecha}.xlsx por WhatsApp
               6. CARLOS  → envía AVANCE_{fecha}.xlsx por WhatsApp
               7. ITALO   → sube hojas a Google Sheets + notifica por WhatsApp
>>>>>>> 3dee96f (Mejoras generales: README detallado, fix CopyPicture, proceso Carlos y consolidacion de pruebas)
```

---

## Arquitectura general

```
Fuentes de datos
    │
    ├── SQL Server (base de datos corporativa de altas)
    ├── rh.xlsx          (mapa vendedor → supervisor → zona)
    ├── lcf.xlsx         (score de riesgo crediticio por vendedor)
    └── cuotas.xlsx      (cuota mensual por vendedor, opcional)
    │
    ▼
AVANCE.py ── ETL principal (~3.500 líneas, 9 secciones numeradas)
    │
    ▼
Archivos Excel generados (en Archivos_Avance/)
    ├── AVANCE_{YYYY-MM-DD}.xlsx          ← dashboard principal multi-hoja
    ├── AVANCE_RESUMIDO.xlsx              ← resumen compacto para Sheets
    └── SEGUIMIENTO_VDD_FIJA_{fecha}.xlsx ← snapshot diario por vendedor
    │
    ▼
main_v2.py ── Orquestador de distribución
    │
    ├── BacksProcess    → Google Sheets (RESUMIDO) + grupo WhatsApp BACKS
    ├── JefesProcess    → captura TDS!B4:V15 y TDS!Y4:AT16 como imágenes
    │                     + SEGUIMIENTO_VDD_FIJA → grupo WhatsApp JEFES
    ├── JesusProcess    → AVANCE_{fecha}.xlsx → WhatsApp contacto Jesús
    ├── CristianProcess → AVANCE_{fecha}.xlsx → WhatsApp contacto Cristian
    ├── GuillermnoProcess→ AVANCE_{fecha}.xlsx → WhatsApp contacto Guillermo
    ├── CarlosProcess   → AVANCE_{fecha}.xlsx → WhatsApp contacto Carlos
    └── ItaloProcess    → hojas MES/DIA → Google Sheets + WhatsApp Italo
            │
            ▼
    wa_client.py (Python)
            │
            ▼ HTTP POST (localhost:8002)
    wa_server.js (Node.js + open-wa)
            │
            ▼
        WhatsApp Web
```

---

## Módulos y responsabilidades

### `AVANCE.py` — ETL principal

Es el núcleo del sistema. Recibe como parámetro el periodo (`YYYY-MM`) y produce todos los archivos Excel del día. Tiene 9 secciones claramente marcadas en el código (ver [Secciones del ETL](#secciones-del-etl-avancepy)).

**Funciones clave:**
- `_pedir_periodo()` — solicita el periodo al usuario si no se pasó por argumento
- `_crear_pivot()` — construye tablas dinámicas usando xlwings (COM de Excel)
- `_color_semaforo()` — aplica rojo (<70%), amarillo (70–90%), verde (>90%) a celdas de avance
- `calc_antiguedad()` — clasifica al vendedor por antigüedad: `<15d`, `>15d`, `>30d`, `>60d`, `>90d`

### `main_v2.py` — Orquestador

Controla el flujo completo: autentica APIs de Google, espera el email trigger, ejecuta `AVANCE.py` como subproceso y luego lanza los 7 procesos de distribución en secuencia.

**Clave:** Marca los emails procesados para que no vuelvan a disparar el pipeline en la siguiente revisión.

### `modules/` — Procesos de distribución

Cada archivo en esta carpeta encapsula la lógica de un destinatario:

| Módulo | Qué hace |
|--------|----------|
| `backs_process.py` | Sube AVANCE_RESUMIDO a Google Sheets; notifica con enlace al grupo BACKS |
| `jefes_process.py` | Abre el AVANCE con xlwings, captura los rangos TDS como imágenes (usando CopyPicture + Pillow), las envía al grupo JEFES junto con el archivo SEGUIMIENTO_VDD_FIJA |
| `jesus_process.py` | Busca AVANCE_{ayer}.xlsx y lo envía a Jesús por WhatsApp |
| `cristian_process.py` | Igual que Jesus, para Cristian |
| `guillermo_process.py` | Igual que Jesus, para Guillermo |
| `carlos_process.py` | Igual que Jesus, para Carlos (primero texto, luego archivo) |
| `italo_process.py` | Sube hojas MES y DIA del AVANCE a Google Sheets; envía enlace a Italo |
| `msg_utils.py` | Función `pick_variant()` para rotar mensajes diariamente |

### `generar_resumido.py` — Resumen por SQL directo

Genera `AVANCE_RESUMIDO.xlsx` ejecutando queries SQL directamente, sin abrir el archivo AVANCE principal con Excel. Útil cuando se quiere regenerar solo el resumen sin correr el ETL completo.

### `whatsapp_server/wa_server.js` — API de WhatsApp

Servidor Express que envuelve la biblioteca `@open-wa/wa-automate`. Implementa una **cola interna con rate limiting** que garantiza mínimo 5 segundos entre cualquier envío, sin importar cuántos procesos Python lo llamen en paralelo.

### `whatsapp_server/wa_client.py` — Cliente HTTP Python

Abstracción Python sobre la API REST del servidor Node.js. Resuelve automáticamente nombres de contactos/grupos desde `config.json`, y reintenta con backoff exponencial (hasta 3 intentos) si el servidor falla.

### `start_wa_server.bat` — Verificación y arranque del servidor

Script `.bat` ejecutado por el Programador de Tareas cada mañana antes del pipeline. Comprueba si el servidor ya está activo (GET /health); si no, lo inicia. Registra el resultado en `logs/wa_server_start.log`.

---

## Estructura del proyecto

```
AVANCE_MOVISTAR/
│
├── AVANCE.py                    # ETL principal (9 secciones, ~3.500 líneas)
├── main_v2.py                   # Orquestador: trigger email → ETL → distribución
├── generar_resumido.py          # Genera AVANCE_RESUMIDO via SQL directo
├── run_modulo.py                # Ejecución manual de un módulo individual
├── start_wa_server.bat          # Verifica/arranca wa_server.js cada mañana
│
├── config.json                  # IDs de WhatsApp, mensajes, rutas (NO en git)
├── config.example.json          # Plantilla de configuración sin datos reales
├── requirements.txt             # Dependencias Python
│
├── modules/                     # Un archivo por proceso de distribución
│   ├── backs_process.py
│   ├── jefes_process.py
│   ├── jesus_process.py
│   ├── cristian_process.py
│   ├── guillermo_process.py
│   ├── carlos_process.py
│   ├── italo_process.py
│   └── msg_utils.py             # Rotación diaria de variantes de mensajes
│
├── whatsapp_server/
│   ├── wa_server.js             # API Express + open-wa (puerto 8002)
│   ├── wa_client.py             # Cliente HTTP Python
│   ├── package.json
│   └── session_data/            # Sesión WhatsApp persistida (NO en git)
│
├── Archivos_Avance/             # Directorio de salida de los Excel (NO en git)
│   ├── AVANCE_{fecha}.xlsx
│   ├── AVANCE_RESUMIDO.xlsx
│   └── SEGUIMIENTO_VDD_FIJA_{fecha}.xlsx
│
├── temp/                        # Imágenes temporales para JefesProcess
└── logs/                        # Logs de ejecución diarios
    ├── automation_{fecha}.log   # Log del orquestador y procesos
    └── wa_server_start.log      # Log de arranques del servidor WhatsApp
```

---

## Instalación desde cero

### Prerrequisitos del sistema

- **Windows 10/11** (requerido para xlwings COM y Programador de Tareas)
- **Microsoft Excel** instalado (xlwings lo controla via COM)
- **Python 3.10+** — se recomienda usar un entorno virtual para aislar dependencias
- **Node.js 18+** — para el servidor WhatsApp
- **ODBC Driver 17 for SQL Server** — para la conexión a la base de datos
- Acceso de red al servidor SQL corporativo
- Cuenta de Google con las APIs de Gmail, Sheets y Drive habilitadas
- WhatsApp activo en el teléfono (para escanear el QR de open-wa)

### Paso 1 — Clonar el repositorio

```bash
git clone https://github.com/BlackZero2x/avance-movistar.git
cd avance-movistar
```

### Paso 2 — Crear entorno virtual Python (recomendado)

Usar un entorno virtual evita que las actualizaciones del sistema rompan las dependencias del proyecto.

```bash
python -m venv venv
venv\Scripts\activate       # activar en Windows
pip install -r requirements.txt
```

Para activar el entorno en sesiones futuras:
```bash
venv\Scripts\activate
```

### Paso 3 — Instalar dependencias Node.js

```bash
cd whatsapp_server
npm install
cd ..
```

### Paso 4 — Configurar credenciales de Google

1. Ir a [Google Cloud Console](https://console.cloud.google.com)
2. Crear un proyecto y habilitar las APIs: Gmail API, Google Sheets API, Google Drive API
3. Crear credenciales OAuth 2.0 → descargar `credentials.json`
4. Colocar `credentials.json` en la raíz del proyecto
5. En la primera ejecución se abrirá una ventana del navegador para autorizar el acceso

### Paso 5 — Crear config.json

```bash
copy config.example.json config.json
```

Editar `config.json` con los datos reales:
- IDs de WhatsApp de grupos y contactos (formato `XXXXXXXXXXX@c.us` o `XXXXXXX-XXXXXXX@g.us`)
- Rutas de directorios locales
- Mensajes y variantes de mensajes

### Paso 6 — Iniciar el servidor WhatsApp por primera vez

```bash
node whatsapp_server/wa_server.js
```

Escaneará un código QR con el teléfono (igual que WhatsApp Web). La sesión se guarda en `whatsapp_server/session_data/` y no necesita re-escanearse a menos que se cierre sesión.

### Paso 7 — Verificar que todo funciona

```bash
# Verificar servidor WhatsApp
python whatsapp_server/wa_client.py --health

# Listar grupos disponibles
python whatsapp_server/wa_client.py --list-groups

# Ejecutar el ETL manualmente (pedirá el periodo YYYY-MM)
python AVANCE.py
```

### Paso 8 — Configurar el Programador de Tareas de Windows

Crear dos tareas programadas:

**Tarea 1: Arranque del servidor WhatsApp**
- Disparador: cada día hábil a las 8:20 AM
- Acción: ejecutar `start_wa_server.bat`
- Propósito: asegurarse de que el servidor WhatsApp está activo antes del pipeline

**Tarea 2: Pipeline principal**
- Disparador: cada día hábil a las 8:30 AM
- Acción: `python C:\AVANCE_MOVISTAR\main_v2.py`
- Marcar "Ejecutar tanto si el usuario inició sesión como si no"
- Marcar "Ejecutar con los privilegios más altos"

---

## Configuración

El archivo `config.json` (no incluido en el repositorio por seguridad) contiene:

```json
{
  "archivos_avance_dir": "C:/AVANCE_MOVISTAR/Archivos_Avance",
  "temp_dir": "C:/AVANCE_MOVISTAR/temp",

  "backs_wa_group": "NOMBRE_O_ID_DEL_GRUPO_BACKS",
  "jefes_wa_group": "NOMBRE_O_ID_DEL_GRUPO_JEFES",

  "jesus_wa_contact": "NOMBRE_O_ID_JESUS",
  "cristian_wa_contact": "NOMBRE_O_ID_CRISTIAN",
  "guillermo_wa_contact": "NOMBRE_O_ID_GUILLERMO",
  "carlos_wa_contact": "NOMBRE_O_ID_CARLOS",
  "italo_wa_contact": "NOMBRE_O_ID_ITALO",

  "carlos_message": "Buenos días, adjunto el avance del día.",
  "carlos_message_variants": [
    "Buenos días, te comparto el reporte de avance.",
    "Hola, aquí está el avance del día de hoy.",
    "Buen día, adjunto el avance actualizado."
  ],

  "jefes_tds_rango1": "B4:V15",
  "jefes_tds_rango2": "Y4:AT16",
  "jefes_mensaje_captura1": "...",
  "jefes_mensaje_captura2": "...",
  "jefes_mensaje_seguimiento": "..."
}
```

Los contactos pueden especificarse por nombre (el cliente resuelve el ID desde `config.json`) o directamente como ID de WhatsApp (`51XXXXXXXXX@c.us` para personas, `XXXXXXXXX-XXXXXXXXX@g.us` para grupos).

---

## Ejecución manual

### Ejecutar el ETL completo

```bash
python AVANCE.py
# Pedirá el periodo: ingresa YYYY-MM (ej: 2026-04)
```

### Ejecutar un módulo de distribución individualmente

```bash
# Ejecuta solo el proceso de un destinatario específico
python run_modulo.py carlos
python run_modulo.py jefes
python run_modulo.py backs
```

### Comandos de prueba del cliente WhatsApp

```bash
# Verificar que el servidor está activo
python whatsapp_server/wa_client.py --health

# Enviar texto de prueba a tu propio número
python whatsapp_server/wa_client.py --test

# Listar todos los grupos
python whatsapp_server/wa_client.py --list-groups

# Buscar un contacto por nombre
python whatsapp_server/wa_client.py --list-contacts "Nombre"

# Probar envío de imagen
python whatsapp_server/wa_client.py --test-image "C:\ruta\imagen.png"

# Probar envío de archivo
python whatsapp_server/wa_client.py --test-file "C:\ruta\archivo.xlsx"
```

### Uso programático desde otro script Python

```python
import sys
sys.path.insert(0, r"C:\AVANCE_MOVISTAR\whatsapp_server")
from wa_client import WhatsAppClient
from msg_utils import pick_variant

wa = WhatsAppClient()

# Enviar texto
wa.send_text("NombreContacto", "Mensaje de prueba")

# Enviar archivo Excel
wa.send_file("NombreGrupo", r"C:\ruta\reporte.xlsx", caption="Reporte del día")

# Enviar imagen
wa.send_image("NombreContacto", r"C:\ruta\captura.png", caption="Dashboard TDS")

# Enviar enlace con texto
wa.send_link("NombreContacto", "https://docs.google.com/...", "Ver en Google Sheets")
```

---

## Secciones del ETL (AVANCE.py)

El script está dividido en 9 secciones claramente delimitadas con encabezados `══` en el código:

| # | Sección | Descripción detallada |
|---|---------|----------------------|
| 1 | **Carga de fuentes** | Conecta a SQL Server y ejecuta una query con CTEs que calcula el periodo actual y anterior. El periodo se controla con `@periodo = 'YYYY-MM'`. También carga los archivos Excel de apoyo (RH, LCF, cuotas). |
| 2 | **Limpieza de fuentes** | Elimina duplicados, convierte tipos de datos, normaliza campos de texto (mayúsculas, quitar espacios), gestiona nulos en campos críticos. |
| 3 | **Join doble con AppVentory** | Enriquece los registros con datos del sistema de ventas AppVentory. Primero intenta cruzar por `codigo_fe`; los registros que no matchean intentan cruzar por `numero_peticion` como segunda oportunidad. Conserva el registro más reciente por `fecha_registro`. |
| 4 | **Join con RH** | Agrega a cada registro el nombre del vendedor, supervisor y zona usando el DNI como clave de join. |
| 5 | **Columnas calculadas** | Genera campos derivados: `DNI_VENDEDOR` (con COALESCE entre múltiples fuentes), clasificación TV, flag `MATCH_DIRECC`, clasificación de antigüedad del vendedor, número de semana. |
| 6 | **Renombres finales** | Estandariza todos los nombres de columnas a los nombres definitivos que aparecerán en el Excel de salida. |
| 7 | **Join LCF → RIESG** | Agrega el score de riesgo crediticio (`RIESG`) por vendedor desde la fuente LCF. |
| 8 | **Limpieza final** | Deduplicación final del dataset consolidado; validación de campos obligatorios; reemplazo de nulos residuales. |
| 9 | **Generación del libro Excel** | Abre Excel via COM (xlwings), construye todas las hojas (TDS, VDD1/2/3, MOVISTAR, MiFibra, SEGUIMIENTO), aplica semáforo de colores, crea tablas dinámicas y aplica formatos visuales. |

**Solo se procesan** registros con `categoria_producto = 'ALTA'` y `fecha_alta IS NOT NULL`.

---

## Salidas que genera el sistema

### AVANCE_{YYYY-MM-DD}.xlsx — Dashboard principal

Libro Excel multi-hoja con toda la información del día:

| Hoja | Contenido |
|------|-----------|
| **TDS** | Dashboard de KPIs: avance vs cuota por zona/supervisor, semáforo de colores, resumen general del periodo |
| **VDD1 / VDD2 / VDD3** | Detalle por vendedor con fórmulas Excel: ventas, cuota asignada, % avance, antigüedad, score de riesgo |
| **MOVISTAR** | Datos crudos de altas de la línea MOVISTAR |
| **MiFibra** | Datos crudos de altas de MiFibra |

### AVANCE_RESUMIDO.xlsx — Resumen para Google Sheets

Versión compacta generada sin abrir Excel (puro Python/openpyxl). Se sube a Google Sheets para que el equipo acceda en tiempo real desde cualquier dispositivo.

### SEGUIMIENTO_VDD_FIJA_{DD-MM-YYYY}.xlsx — Snapshot diario

Copia de las hojas VDD del día, pensada para acumularse y comparar evolución de vendedores semana a semana. Se envía al grupo de jefes.

---

## Servidor y cliente de WhatsApp

### Endpoints disponibles (wa_server.js)

| Endpoint | Método | Parámetros | Descripción |
|----------|--------|-----------|-------------|
| `/health` | GET | — | Retorna `{ status: "ok" }` si el servidor está listo |
| `/list-groups` | GET | — | Lista todos los grupos de WhatsApp con nombre e ID |
| `/list-contacts` | GET | `?name=texto` | Busca contactos que coincidan con el nombre |
| `/send-text` | POST | `{ to, message }` | Envía mensaje de texto (encolado) |
| `/send-image` | POST | `{ to, path, caption }` | Envía imagen desde ruta local (convertida a base64, encolada) |
| `/send-file` | POST | `{ to, path, caption }` | Envía archivo (xlsx, pdf, csv…) desde ruta local (encolado) |
| `/send-mention` | POST | `{ to, message, mentions }` | Envía texto con menciones (@usuario) en grupos (encolado) |
| `/send-link` | POST | `{ to, url, text }` | Envía texto con previsualización de enlace (encolado) |

Todos los envíos pasan por una **cola interna** que inserta automáticamente un delay mínimo de 5 segundos entre mensajes para evitar el spam detector de WhatsApp.

### Cómo resuelve destinatarios wa_client.py

El campo `to` puede ser:
- Un **nombre de clave** del `config.json` (ej: `"jefes_wa_group"`) — el cliente resuelve el ID automáticamente
- Un **nombre parcial** (ej: `"Canal Fija 2026"`) — el cliente busca el ID en el config
- Un **ID directo de WhatsApp** (ej: `51XXXXXXXXX@c.us`) — se usa tal cual

---

## Mitigación de riesgo de ban en WhatsApp

WhatsApp detecta y penaliza cuentas que envían mensajes masivos o repetitivos en ráfagas. Este sistema implementa tres capas de protección:

### 1. Cola con rate limiting (wa_server.js)
Todos los envíos pasan por una cola FIFO. El servidor espera mínimo **5 segundos** entre cada salida, sin importar cuántos procesos llamen al servidor simultáneamente. Esto previene ráfagas incluso cuando múltiples módulos Python se ejecutan en cascada.

### 2. Rotación diaria de mensajes (msg_utils.py)
En lugar de enviar exactamente el mismo texto cada día, `pick_variant()` elige una variante diferente según un hash de la fecha. Con 3–4 variantes por destinatario, el texto cambia diariamente sin intervención manual.

```python
from msg_utils import pick_variant

mensaje = pick_variant(
    variants=["Buenos días, adjunto el avance.", "Buen día, aquí el reporte.", "Hola, te comparto el avance de hoy."],
    fallback="Buenos días, adjunto el avance del día."
)
# Devuelve siempre la misma variante para una fecha dada (determinista)
```

### 3. Delays adicionales en procesos de alto volumen
El proceso `JefesProcess` envía varias imágenes seguidas al mismo grupo. Para reducir la tasa de envíos incluso por encima del rate limiting del servidor:
- 12 segundos de espera entre cada imagen TDS
- 10 segundos antes de enviar el archivo SEGUIMIENTO_VDD

---

## Stack tecnológico

| Capa | Tecnología | Por qué se usa |
|------|-----------|----------------|
| ETL / procesamiento de datos | Python 3.10, pandas 2.x | Manipulación eficiente de DataFrames con múltiples joins y transformaciones |
| Generación Excel (lectura/escritura simple) | openpyxl | Crear y modificar archivos .xlsx sin abrir Excel |
| Automatización Excel avanzada | xlwings (COM) | Controlar Excel via COM para tablas dinámicas, semáforo y capturas de pantalla |
| Captura de pantalla desde Excel | Pillow (ImageGrab), win32gui | Capturar rangos de Excel como imágenes PNG para enviarlas por WhatsApp |
| Base de datos | SQL Server 2019 + pyodbc + SQLAlchemy | Fuente principal de registros de altas |
| Servidor WhatsApp | Node.js 18 + @open-wa/wa-automate + Express | API REST local que encapsula la automatización de WhatsApp Web |
| Cliente WhatsApp | Python + requests | Llamadas HTTP al servidor Node.js con reintentos y resolución de contactos |
| Google APIs | google-api-python-client, gspread | Acceso a Gmail (trigger), Sheets (subida de datos) y Drive (gestión archivos) |
| Programación de tareas | Windows Task Scheduler | Disparar el pipeline automáticamente cada mañana |
| Logging | Python logging (RotatingFileHandler) | Registro diario de ejecuciones con timestamps para auditoría y debugging |

---

## Glosario del dominio

| Término | Significado |
|---------|-------------|
| **ALTA** | Cliente que completó el proceso de activación de servicio fijo |
| **TDS** | Resumen Técnico de Datos — hoja principal del dashboard con KPIs agregados |
| **VDD** | Detalle por Vendedor — desglose de rendimiento individual de cada asesor comercial |
| **RH** | Recursos Humanos — tabla que mapea cada vendedor (DNI) a su supervisor y zona |
| **LCF** | Fuente de datos de riesgo crediticio; aporta el campo `RIESG` por vendedor |
| **AVANCE** | Porcentaje de cumplimiento de la cuota mensual (altas reales / cuota asignada × 100) |
| **eAuren** | Nombre de la base de datos en SQL Server donde se registran las altas |
| **FE** | Código interno del vendedor en el sistema AppVentory (usado para joins) |
| **AppVentory** | Sistema corporativo de registro de ventas; fuente de datos de vendedor y petición |
| **SEGUIMIENTO** | Snapshot diario de las hojas VDD, usado para comparar evolución semanal |
| **D-1** | Dato del día anterior (el pipeline trabaja con datos de ayer, disponibles a las 8:30 AM) |
| **BACKS** | Grupo de analistas que reciben el resumen en Google Sheets |
| **Periodo** | Mes de análisis en formato `YYYY-MM`; controla qué datos extrae la query SQL |

---

## Preguntas frecuentes y solución de problemas

### El servidor WhatsApp no arranca

1. Verificar que Node.js está instalado: `node --version`
2. Verificar que las dependencias están instaladas: `cd whatsapp_server && npm install`
3. Si la sesión expiró, borrar `whatsapp_server/session_data/` y volver a escanear el QR
4. Revisar `logs/wa_server_YYYY-MM-DD.log` para el error exacto

### `CopyPicture failed` — las capturas TDS no se generan

Este error ocurre cuando Excel no es la ventana activa al momento de ejecutar `CopyPicture`. Causas comunes:
- El script corre desde el Programador de Tareas en una sesión no interactiva
- Otra ventana tomó el foco entre la apertura de Excel y la captura

La solución implementada en `jefes_process.py` usa `AllowSetForegroundWindow` + `SetForegroundWindow` via ctypes y win32gui. Si persiste el error, asegurarse de que la tarea del Programador de Tareas tiene habilitada la opción **"Ejecutar solo cuando el usuario haya iniciado sesión"** temporalmente para diagnosticar.

### Los paquetes Python no se encuentran después de actualizar Python

Al actualizar Python en Windows, el PATH apunta a la nueva versión que no tiene los paquetes instalados. Solución:

```bash
# Con el entorno virtual activado, reinstalar todo
venv\Scripts\activate
pip install -r requirements.txt
```

Si no se usa entorno virtual, reinstalar en la versión nueva:
```bash
python -m pip install -r requirements.txt
```

### El ETL genera el Excel pero no envía por WhatsApp

1. Verificar que el servidor está activo: `python whatsapp_server/wa_client.py --health`
2. Si el servidor no responde, ejecutar `start_wa_server.bat` manualmente
3. Revisar `logs/automation_{fecha}.log` para ver el error exacto del proceso de distribución

### Un contacto no recibe mensajes ("Not a contact")

open-wa en versión gratuita solo puede enviar mensajes a números que estén guardados como contactos en el teléfono donde corre el servidor. Verificar que:
1. El número está guardado en la agenda del teléfono
2. El servidor WhatsApp fue reiniciado después de agregar el contacto (carga la lista al iniciar)
3. El número está en formato internacional sin el `+` (ej: `51XXXXXXXXX`)

---

## Licencia

Uso interno. El código es de autoría propia. Los datos de producción (config.json, session_data/, Archivos_Avance/) no están incluidos en este repositorio.
