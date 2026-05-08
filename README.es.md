# Reporte y Sistema de Seguimiento de Ventas

> [English version](README.md)

Sistema de automatización end-to-end para el seguimiento diario de ventas de la línea fija de Movistar Perú. Consolida datos de múltiples fuentes, genera dashboards Excel con indicadores de gestión y los distribuye automáticamente a los integrantes del equipo comercial por WhatsApp y correo electrónico, sin intervención manual.

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
- **¿Cómo evoluciona cada vendedor día a día?** (RT, ALTAS y CONSULTAS diarias con alertas)

Este sistema genera esa información de forma automática y la entrega a las personas correctas (jefes de zona, gerentes, analistas) a través de WhatsApp y Google Sheets, **sin que nadie tenga que ejecutar nada manualmente**.

---

## Cómo funciona en el día a día

```
08:30 AM  →  El Programador de Tareas de Windows inicia main_v2.py
             y verifica que el servidor de WhatsApp esté activo.

08:30–??  →  main_v2.py revisa el correo cada 10 minutos,
             buscando el email trigger que indica que los datos
             del día ya están disponibles en el sistema.

Al llegar →  AVANCE.py se ejecuta:
el email       - Consulta SQL Server + CSV local + Google Sheet MiFibra
               - Cruza con archivos Excel de RH, riesgo crediticio y cuotas
               - Calcula KPIs, semáforos, tablas dinámicas y seguimiento diario
               - Genera los archivos Excel finales

Después  →   7 procesos de distribución corren en cadena:
               1. BACKS    → sube resumen a Google Sheets + notifica al equipo
               2. JEFES    → captura imágenes TDS + envía SEGUIMIENTO al grupo
               3. JESUS    → envía AVANCE_{fecha}.xlsx por WhatsApp
               4. CRISTIAN → envía AVANCE_{fecha}.xlsx por WhatsApp
               5. GUILLERMO→ envía AVANCE_{fecha}.xlsx por WhatsApp y por correo
               6. CARLOS   → envía AVANCE_{fecha}.xlsx solo por correo
               7. ITALO    → sube hojas a Google Sheets + notifica por WhatsApp
```

---

## Arquitectura general

```
Fuentes de datos
    │
    ├── SQL Server (eAuren) — altas, registros totales, consultas DITO
    ├── rh.xlsx             — mapa vendedor → supervisor → zona
    ├── lcf.xlsx            — score de riesgo crediticio
    ├── cuotas.xlsx / cuotas_zonal_sup.xlsx  — cuotas por vendedor y supervisor
    ├── BD_Ventas_AUREN.csv — ventas MiFibra (MF_CSV_PATH en .env)
    └── Google Sheet privado — ventas MiFibra adicionales (MF_SHEET_ID en .env)
    │
    ▼
AVANCE.py — ETL principal (~3.500 líneas, 9 secciones numeradas)
    │       Incluye helpers de seguimiento diario (integrados, sin módulo externo)
    │
    ▼
Archivos Excel generados (en Archivos_Avance/)
    ├── AVANCE_{YYYY-MM-DD}.xlsx           ← dashboard principal multi-hoja
    ├── AVANCE_RESUMIDO.xlsx               ← resumen compacto para Sheets
    └── SEGUIMIENTO_VDD_FIJA_{fecha}.xlsx  ← VDD1 + VDD2 + VDD3 del día
    │
    ▼
main_v2.py — Orquestador de distribución
    │
    ├── BacksProcess     → Google Sheets (RESUMIDO) + grupo WhatsApp BACKS
    ├── JefesProcess     → captura TDS!B4:V15 y TDS!Y4:AT17 como imágenes
    │                      + SEGUIMIENTO_VDD_FIJA → grupo WhatsApp JEFES
    ├── JesusProcess     → AVANCE_{fecha}.xlsx → WhatsApp contacto Jesús
    ├── CristianProcess  → AVANCE_{fecha}.xlsx → WhatsApp contacto Cristian
    ├── GuillermnoProcess→ AVANCE_{fecha}.xlsx → WhatsApp + correo Guillermo
    ├── CarlosProcess    → AVANCE_{fecha}.xlsx → correo Carlos (solo email)
    └── ItaloProcess     → hojas MES/DIA → Google Sheets + WhatsApp Italo
            │
            ▼
    wa_client.py (Python)
            │
            ▼ HTTP POST (localhost:8002)
    wa_server.js (Node.js + whatsapp-web.js)
            │
            ▼
        WhatsApp Web
```

---

## Módulos y responsabilidades

### `AVANCE.py` — ETL principal

Núcleo del sistema. Recibe como parámetro el periodo (`YYYY-MM`) y produce todos los archivos Excel del día. Tiene 9 secciones claramente marcadas en el código.

Incluye directamente los helpers de seguimiento diario (función `_seg_agregar_hoja`), por lo que no depende de ningún módulo externo para generar la hoja VDD2.

**Funciones clave:**
- `_pedir_periodo()` — solicita el periodo si no se pasó por argumento
- `_crear_pivot()` — construye tablas dinámicas via xlwings (COM de Excel)
- `_color_semaforo()` — aplica rojo (<70%), amarillo (70–90%), verde (>90%)
- `calc_antiguedad()` — clasifica al vendedor: `<15d`, `>15d`, `>30d`, `>60d`, `>90d`
- `_normalizar_zonal2()` — deriva el campo `zonal2`: sub-zonales LIMA* → `"LIMA"`, resto igual a `zonal`
- `_seg_agregar_hoja()` — genera la hoja VDD2 (seguimiento diario) en cualquier workbook openpyxl

### `main_v2.py` — Orquestador

Controla el flujo completo: autentica APIs de Google, espera el email trigger, ejecuta `AVANCE.py` como subproceso y luego lanza los 7 procesos de distribución en secuencia. Marca los emails procesados para no volver a disparar el pipeline.

### `modules/` — Procesos de distribución

| Módulo | Canal | Qué hace |
|--------|-------|----------|
| `backs_process.py` | WhatsApp | Sube AVANCE_RESUMIDO a Google Sheets; notifica al grupo BACKS con enlace |
| `jefes_process.py` | WhatsApp | Captura rangos TDS (`B4:V15` y `Y4:AT17`) como imágenes con CopyPicture + Pillow; las envía junto con SEGUIMIENTO_VDD_FIJA al grupo JEFES |
| `jesus_process.py` | WhatsApp | Envía AVANCE_{ayer}.xlsx a Jesús |
| `cristian_process.py` | WhatsApp | Envía AVANCE_{ayer}.xlsx a Cristian |
| `guillermo_process.py` | WhatsApp + correo | Envía AVANCE_{ayer}.xlsx a Guillermo por WhatsApp y por Gmail a `guillermoj.hinostroza@auren.com.pe` |
| `carlos_process.py` | Correo | Envía AVANCE_{ayer}.xlsx solo por Gmail a `carlos.parra@auren.com.pe` |
| `italo_process.py` | WhatsApp | Sube hojas MES y DIA del AVANCE a Google Sheets; envía enlace a Italo |
| `shared/gmail_helper.py` | — | Envío de correos con adjuntos via Gmail API (cuenta `augusto.moreno@auren.com.pe`) |
| `msg_utils.py` | — | `pick_variant()`: elige una variante de mensaje por hash de fecha para rotar textos diariamente |

### `generar_resumido.py` — Resumen por SQL directo

Genera `AVANCE_RESUMIDO.xlsx` ejecutando queries SQL directamente, sin abrir el archivo AVANCE con Excel. Las hojas RTCHB/ALTASCHB incluyen las zonales CHIMBOTE, NORTE CHICO y todas las que empiezan con "LIMA".

### `whatsapp_server/wa_server.js` — API de WhatsApp

Servidor Express que envuelve `whatsapp-web.js` (con sesión persistida via `LocalAuth` en `session_data/`). Usa Chrome instalado para Puppeteer. Implementa una **cola con rate limiting** que garantiza mínimo 5 segundos entre envíos. Las menciones se resuelven con `getContactById()` con fallback para números no guardados en agenda.

### `whatsapp_server/wa_client.py` — Cliente HTTP Python

Abstracción Python sobre la API REST del servidor Node.js. Resuelve nombres desde `config.json` y reintenta con backoff exponencial (hasta 3 intentos).

### `start_wa_server.bat` — Verificación y arranque del servidor

Ejecutado por el Programador de Tareas cada mañana. Comprueba si el servidor está activo (GET /health); si no, lo inicia. Al arrancar, envía notificación WhatsApp confirmando que el servidor está activo.

---

## Estructura del proyecto

```
AVANCE_MOVISTAR/
│
├── AVANCE.py                    # ETL principal + helpers VDD2 integrados (~3.500 líneas)
├── main_v2.py                   # Orquestador: trigger email → ETL → distribución
├── generar_resumido.py          # Genera AVANCE_RESUMIDO via SQL directo
├── generar_seguimiento_diario.py# Script standalone para VDD2 (uso manual/debug)
├── run_modulo.py                # Ejecución manual de un módulo individual
├── start_wa_server.bat          # Verifica/arranca wa_server.js cada mañana
├── stop_wa_server.bat           # Detiene wa_server.js
│
├── .env                         # Variables de entorno (NO en git)
├── config.json                  # IDs de WhatsApp, mensajes, rutas (NO en git)
├── config.example.json          # Plantilla de configuración sin datos reales
├── credentials.json             # OAuth Google (NO en git)
├── token.json                   # Token OAuth Google (NO en git)
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
│   ├── msg_utils.py             # Rotación diaria de variantes de mensajes
│   └── shared/
│       └── gmail_helper.py      # Envío de correos via Gmail API
│
├── whatsapp_server/
│   ├── wa_server.js             # API Express + whatsapp-web.js (puerto 8002)
│   ├── wa_client.py             # Cliente HTTP Python
│   ├── package.json
│   ├── config.json              # IDs resueltos por el servidor
│   ├── session_data/            # Sesión WhatsApp persistida (NO en git)
│   └── logs/                    # Logs diarios del servidor WA
│
├── Archivos_Avance/             # Directorio de salida de los Excel (NO en git)
│   ├── AVANCE_{fecha}.xlsx
│   ├── AVANCE_RESUMIDO.xlsx
│   └── SEGUIMIENTO_VDD_FIJA_{fecha}.xlsx
│
└── temp/                        # Imágenes temporales para JefesProcess
```

---

## Instalación desde cero

### Prerrequisitos del sistema

- **Windows 10/11** (requerido para xlwings COM y Programador de Tareas)
- **Microsoft Excel** instalado (xlwings lo controla via COM)
- **Python 3.10+**
- **Node.js 18+**
- **Google Chrome** instalado en `C:\Program Files\Google\Chrome\Application\chrome.exe`
- **ODBC Driver 17 for SQL Server**
- Acceso de red al servidor SQL corporativo (`AUREN22\AUREN`)
- Cuenta de Google con las APIs de Gmail, Sheets y Drive habilitadas

### Paso 1 — Clonar el repositorio

```bash
git clone https://github.com/BlackZero2x/avance_ventas_telcom.git
cd avance_ventas_telcom
```

### Paso 2 — Crear entorno virtual Python

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### Paso 3 — Instalar dependencias Node.js

```bash
cd whatsapp_server
npm install
cd ..
```

### Paso 4 — Configurar credenciales de Google

1. Ir a [Google Cloud Console](https://console.cloud.google.com)
2. Crear un proyecto y habilitar: Gmail API, Google Sheets API, Google Drive API
3. Crear credenciales OAuth 2.0 → descargar como `credentials.json`
4. Colocar `credentials.json` en la raíz del proyecto
5. En la primera ejecución se abrirá el navegador para autorizar; esto genera `token.json`

### Paso 5 — Crear archivos de configuración

```bash
copy config.example.json config.json
```

Crear el archivo `.env` en la raíz con las variables requeridas (ver sección [Configuración](#configuración)).

### Paso 6 — Iniciar el servidor WhatsApp por primera vez

```bash
node whatsapp_server/wa_server.js
```

Escaneará un código QR con el teléfono (igual que WhatsApp Web). La sesión se guarda en `whatsapp_server/session_data/` y no necesita re-escanearse.

### Paso 7 — Verificar que todo funciona

```bash
# Verificar servidor WhatsApp
python whatsapp_server/wa_client.py --health

# Listar grupos disponibles
python whatsapp_server/wa_client.py --list-groups

# Ejecutar el ETL manualmente
python AVANCE.py
```

### Paso 8 — Configurar el Programador de Tareas de Windows

**Tarea 1: Arranque del servidor WhatsApp**
- Disparador: cada día hábil a las 8:20 AM
- Acción: ejecutar `start_wa_server.bat`

**Tarea 2: Pipeline principal**
- Disparador: cada día hábil a las 8:30 AM
- Acción: `python C:\proyectos\AVANCE_MOVISTAR\main_v2.py`
- Marcar "Ejecutar tanto si el usuario inició sesión como si no"
- Marcar "Ejecutar con los privilegios más altos"

---

## Configuración

### Variables de entorno (.env)

| Variable | Descripción |
|----------|-------------|
| `SQL_SERVER` | Servidor SQL (`AUREN22\AUREN`) |
| `SQL_DATABASE` | Base de datos (`eAuren`) |
| `SQL_USER` / `SQL_PASSWORD` | Credenciales SQL |
| `MF_CSV_PATH` | Ruta al CSV local `BD_Ventas_AUREN.csv` |
| `MF_SHEET_ID` | ID del Google Sheet privado de MiFibra |
| `URL_VENTORY` / `URL_RH` / `URL_LCF` | URLs CSV públicas de Google Sheets |
| `HTTP_PROXY` / `HTTPS_PROXY` / `NO_PROXY` | Proxy corporativo (si aplica) |
| `AVANCE_DIR` | Ruta raíz del proyecto |
| `CHECK_INTERVAL_MINUTES` | Intervalo de polling del orquestador (default 10) |

### config.json

```json
{
  "groups": {
    "Canal Fija 2026": "XXXXXXXX-XXXXXXXX@g.us",
    "JEFES": "XXXXXXXX-XXXXXXXX@g.us"
  },
  "contacts": {
    "Jesús": "51XXXXXXXXX@c.us",
    "Cristian": "51XXXXXXXXX@c.us"
  },
  "jefes_tds_rango1": "B4:V15",
  "jefes_tds_rango2": "Y4:AT17"
}
```

---

## Ejecución manual

### Ejecutar el ETL completo

```bash
python AVANCE.py
# Pedirá el periodo: ingresa YYYY-MM (ej: 2026-05)
```

### Ejecutar un módulo de distribución individualmente

```bash
python run_modulo.py carlos
python run_modulo.py jefes
python run_modulo.py backs
```

### Comandos de prueba del cliente WhatsApp

```bash
python whatsapp_server/wa_client.py --health
python whatsapp_server/wa_client.py --test
python whatsapp_server/wa_client.py --list-groups
python whatsapp_server/wa_client.py --list-contacts "Nombre"
python whatsapp_server/wa_client.py --test-image "C:\ruta\imagen.png"
python whatsapp_server/wa_client.py --test-file  "C:\ruta\archivo.xlsx"
python whatsapp_server/wa_client.py --test-link  "https://..."
```

### Uso programático desde Python

```python
import sys
sys.path.insert(0, r"C:\proyectos\AVANCE_MOVISTAR\whatsapp_server")
from wa_client import WhatsAppClient
from msg_utils import pick_variant

wa = WhatsAppClient()
wa.send_text("Canal Fija 2026", "Mensaje")
wa.send_file("Back de AUREN 2025", r"C:\ruta\reporte.xlsx", caption="Reporte")
wa.send_image("Cristian", r"C:\ruta\captura.png", caption="TDS del día")
wa.send_link("Jesús", "https://drive.google.com/...", "Seguimiento FIJA")
wa.send_mention("Canal Fija 2026", "Hola @número", ["51962969371@c.us"])
```

---

## Secciones del ETL (AVANCE.py)

| # | Sección | Descripción |
|---|---------|-------------|
| 1 | **Carga de fuentes** | SQL Server (altas, RT, consultas DITO, mes anterior) + RH, LCF, cuotas en Excel + CSV MiFibra + Google Sheet MiFibra privado |
| 2 | **Limpieza de fuentes** | Deduplicación, conversión de tipos, normalización de texto, gestión de nulos |
| 3 | **Join doble con AppVentory** | Match por `codigo_fe`; fallback por `numero_peticion`; conserva registro más reciente |
| 4 | **Join con RH** | Agrega nombre del vendedor, supervisor y zona por DNI |
| 5 | **Columnas calculadas** | `DNI_VENDEDOR` (COALESCE entre fuentes), clasificación TV, `MATCH_DIRECC`, antigüedad, semana |
| 6 | **Renombres finales** | Estandariza nombres de columnas al esquema de salida |
| 7 | **Join LCF → RIESG** | Agrega score de riesgo crediticio por vendedor |
| 8 | **Limpieza final** | Deduplicación final, validación de campos obligatorios |
| 9 | **Generación del libro Excel** | TDS (tablas 1 y 2, semáforo), VDD1/VDD2/VDD3, MOVISTAR (pivots diarios por `zonal2`), MiFibra, hojas RT/ALTAS con campo `zonal2`, SEGUIMIENTO_VDD_FIJA |

Solo se procesan registros con `categoria_producto = 'ALTA'` y `fecha_alta IS NOT NULL`.

---

## Salidas que genera el sistema

### AVANCE_{YYYY-MM-DD}.xlsx — Dashboard principal

| Hoja | Contenido |
|------|-----------|
| **MOVISTAR** | Pivots diarios de conversión, altas (total/regular/flex), velocidades y VDD por `zonal2` |
| **MiFibra** | Dashboard de ventas e instalaciones MiFibra por filial y plan |
| **TDS** | KPIs por zonal (tabla 1, cols B–V) y por supervisor (tabla 2, cols Y–AT), semáforo rojo/amarillo/verde |
| **VDD1** | Detalle por vendedor con fórmulas Excel |
| **VDD2** | Seguimiento diario: RT, ALTAS y CONSULTAS por día + resumen últimos 3 días + columna ALERTAS |
| **VDD3** | Tablas dinámicas nativas |
| **RT** | Registros totales crudos con campo `zonal2` |
| **ALTAS** | Altas crudas con campo `zonal2` |
| **RH**, **VENTORY**, **CON** | Fuentes auxiliares |
| **MES**, **DIA** | Datos para Italo (ocultas) |

### SEGUIMIENTO_VDD_FIJA_{DD-MM-YYYY}.xlsx

Copia de VDD1, VDD2 y VDD3 del día. Se envía al grupo de jefes.

### AVANCE_RESUMIDO.xlsx

Versión compacta generada con SQL directo (sin abrir Excel). Se sube a Google Sheets.

---

## Servidor y cliente de WhatsApp

### Endpoints disponibles (wa_server.js, puerto 8002)

| Endpoint | Método | Descripción |
|----------|--------|-------------|
| `/health` | GET | Estado del servidor (`ready` / `not_ready`) |
| `/list-groups` | GET | Lista todos los grupos con nombre, ID y número de participantes |
| `/list-contacts?name=` | GET | Busca contactos por nombre |
| `/send-text` | POST | Envía texto (encolado, ≥5 s entre mensajes) |
| `/send-image` | POST | Envía imagen desde ruta local (encolado) |
| `/send-file` | POST | Envía archivo desde ruta local (encolado) |
| `/send-mention` | POST | Envía texto con menciones; resuelve contactos con `getContactById()` + fallback para no-agenda |
| `/send-link` | POST | Envía texto + URL (encolado) |

### Resolución de destinatarios

El campo `to` puede ser un nombre de clave en `config.json` o un ID directo de WhatsApp (`51XXXXXXXXX@c.us` para personas, `XXXXXXXX-XXXXXXXX@g.us` para grupos).

---

## Mitigación de riesgo de ban en WhatsApp

### 1. Cola con rate limiting (wa_server.js)
Cola FIFO con mínimo **5 segundos** entre envíos, sin importar cuántos módulos llamen en paralelo.

### 2. Rotación diaria de mensajes (msg_utils.py)
`pick_variant()` elige una variante distinta por hash de fecha. Con 3–4 variantes, el texto cambia diariamente de forma determinista.

```python
from modules.msg_utils import pick_variant

mensaje = pick_variant(
    variants=["Buenos días, adjunto el avance.", "Buen día, aquí el reporte.", "Hola, te comparto el avance."],
    fallback="Buenos días, adjunto el avance del día."
)
```

### 3. Delays adicionales en JefesProcess
- 12 segundos entre imágenes TDS
- 10 segundos antes del archivo SEGUIMIENTO_VDD_FIJA

---

## Stack tecnológico

| Capa | Tecnología | Uso |
|------|-----------|-----|
| ETL / datos | Python 3.10, pandas 2.x | Joins, transformaciones, pivots |
| Excel (escritura simple) | openpyxl | Crear/modificar .xlsx, hoja VDD2 |
| Excel avanzado | xlwings (COM) | Tablas dinámicas, semáforo, capturas de pantalla |
| Capturas desde Excel | Pillow, win32gui | Rangos TDS → imágenes PNG |
| Base de datos | SQL Server + pyodbc + SQLAlchemy | Fuente principal de registros |
| Servidor WhatsApp | Node.js 18 + whatsapp-web.js + Express | API REST local sobre WhatsApp Web |
| Cliente WhatsApp | Python + requests | HTTP al servidor Node con reintentos |
| Google APIs | google-api-python-client | Gmail (trigger), Sheets (datos), Drive |
| Correo electrónico | Gmail API | Envío a Guillermo y Carlos |
| Programación de tareas | Windows Task Scheduler | Disparo automático diario |

---

## Glosario del dominio

| Término | Significado |
|---------|-------------|
| **ALTA** | Cliente que completó el proceso de activación de servicio fijo |
| **ALTAS.MF** | Instalaciones MiFibra del período (CSV local + Google Sheet privado, deduplicadas por DNI + fecha) |
| **RT** | Registros Totales — todos los registros, no solo ALTAS |
| **CON** | Consultas DITO (intenciones de compra) |
| **TDS** | Resumen Técnico de Datos — dashboard principal con KPIs agregados |
| **VDD** | Detalle por Vendedor — desglose de rendimiento individual |
| **VDD2** | Seguimiento diario: RT, ALTAS y CON por día + alertas automáticas |
| **zonal2** | Campo normalizado: sub-zonales LIMA* → `"LIMA"`; resto igual a `zonal` |
| **RH** | Recursos Humanos — mapeo vendedor (DNI) → supervisor → zona |
| **LCF** | Fuente de riesgo crediticio — aporta el campo `RIESG` |
| **AVANCE** | Porcentaje de cumplimiento de cuota (altas / cuota × 100) |
| **eAuren** | Base de datos SQL Server donde se registran las altas |
| **FE** | Código interno del vendedor en AppVentory (clave de join) |
| **AppVentory** | Sistema corporativo de registro de ventas |
| **BACKS** | Grupo de analistas que reciben el resumen en Google Sheets |
| **Periodo** | Mes de análisis en formato `YYYY-MM` |
| **D-1** | Los datos disponibles cada mañana corresponden al día anterior |

---

## Preguntas frecuentes y solución de problemas

### El servidor WhatsApp no arranca

1. Verificar que Node.js está instalado: `node --version`
2. Verificar dependencias: `cd whatsapp_server && npm install`
3. Si la sesión expiró, borrar `whatsapp_server/session_data/` y re-escanear el QR
4. Revisar `whatsapp_server/logs/wa_server_YYYY-MM-DD.log`

### Las menciones aparecen como texto literal (`@51XXXXXXXXX`)

El servidor usa `whatsapp-web.js` y requiere objetos `Contact` en el parámetro `mentions`, no strings. El endpoint `/send-mention` resuelve cada ID con `getContactById()`; si el número no está en agenda del teléfono, usa un objeto de fallback mínimo para que la mención se renderice de todas formas. Si el problema persiste, reiniciar `wa_server.js` para que recargue la lista de contactos.

### `CopyPicture failed` — las capturas TDS no se generan

Ocurre cuando Excel no es la ventana activa. `jefes_process.py` usa `AllowSetForegroundWindow` + `SetForegroundWindow` para mitigarlo. Si persiste, asegurarse de que la tarea del Programador de Tareas tiene habilitada la opción **"Ejecutar solo cuando el usuario haya iniciado sesión"** para diagnosticar.

### LIMA no aparece en la hoja MOVISTAR

Los pivots de MOVISTAR usan el campo `zonal2`. Verificar que `altas_df` y `rt_df` tienen el campo `zonal` poblado antes de que `_normalizar_zonal2()` lo procese.

### El ETL genera el Excel pero no envía por WhatsApp

1. Verificar servidor: `python whatsapp_server/wa_client.py --health`
2. Si no responde, ejecutar `start_wa_server.bat` manualmente
3. Revisar logs del orquestador

### MiFibra no carga datos del Google Sheet

El proceso continúa solo con el CSV local si el Sheet falla (token vencido, sin red). Verificar que `token.json` está vigente ejecutando cualquier script que use la API de Google (se refresca automáticamente si el refresh token sigue válido).

---

## Licencia

Uso interno. El código es de autoría propia. Los datos de producción (`.env`, `config.json`, `session_data/`, `Archivos_Avance/`, `credentials.json`, `token.json`) no están incluidos en este repositorio.
