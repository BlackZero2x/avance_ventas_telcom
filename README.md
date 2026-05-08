# Sales Tracking & Reporting System

> [Versión en español](README.es.md)

End-to-end automation system for daily fixed-line sales tracking at Movistar Peru. Consolidates data from multiple sources, generates Excel dashboards with management KPIs, and automatically distributes reports to the commercial team via WhatsApp and email — no manual intervention required.

---

## Table of Contents

1. [What this system does](#what-this-system-does)
2. [How it works day to day](#how-it-works-day-to-day)
3. [Architecture](#architecture)
4. [Modules and responsibilities](#modules-and-responsibilities)
5. [Project structure](#project-structure)
6. [Installation](#installation)
7. [Configuration](#configuration)
8. [Manual execution](#manual-execution)
9. [ETL sections (AVANCE.py)](#etl-sections-avancepy)
10. [System outputs](#system-outputs)
11. [WhatsApp server and client](#whatsapp-server-and-client)
12. [WhatsApp ban-risk mitigation](#whatsapp-ban-risk-mitigation)
13. [Tech stack](#tech-stack)
14. [Domain glossary](#domain-glossary)
15. [Troubleshooting](#troubleshooting)

---

## What this system does

Every business day, the fixed-line sales team needs to know:

- **How many activations were recorded?** (customers activated that day)
- **What is the progress vs quota?** (fulfillment percentage by supervisor/zone)
- **Which salespeople are at risk?** (low performance flagged with red/yellow/green traffic lights)
- **What is each salesperson's daily breakdown?** (RT, activations, and consultations per day with automatic alerts)

This system generates that information automatically and delivers it to the right people (zone managers, executives, analysts) through WhatsApp and Google Sheets.

---

## How it works day to day

```
08:30 AM  →  Windows Task Scheduler starts main_v2.py
             and checks that the WhatsApp server is active.

08:30–??  →  main_v2.py polls the inbox every 10 minutes,
             looking for the trigger email that signals
             the day's data is available in the system.

On arrival→  AVANCE.py runs:
             - Queries SQL Server + local CSV + private Google Sheet (MiFibra)
             - Joins with HR, credit risk, and quota Excel files
             - Calculates KPIs, traffic lights, pivot tables, and daily tracking
             - Produces the final Excel files

After    →   7 distribution processes run in sequence:
               1. BACKS    → uploads summary to Google Sheets + notifies team
               2. JEFES    → captures TDS dashboard screenshots + sends report
               3. JESUS    → sends AVANCE_{date}.xlsx via WhatsApp
               4. CRISTIAN → sends AVANCE_{date}.xlsx via WhatsApp
               5. GUILLERMO→ sends AVANCE_{date}.xlsx via WhatsApp + email
               6. CARLOS   → sends AVANCE_{date}.xlsx via email only
               7. ITALO    → uploads sheets to Google Sheets + notifies via WhatsApp
```

---

## Architecture

```
Data sources
    │
    ├── SQL Server (eAuren)  — activations, total records, DITO consultations
    ├── rh.xlsx              — salesperson → supervisor → zone mapping
    ├── lcf.xlsx             — credit risk scores
    ├── cuotas*.xlsx         — monthly quotas by salesperson and supervisor
    ├── BD_Ventas_AUREN.csv  — MiFibra sales (MF_CSV_PATH in .env)
    └── Private Google Sheet — additional MiFibra sales (MF_SHEET_ID in .env)
    │
    ▼
AVANCE.py  —  Main ETL (~3,500 lines, 9 numbered sections)
    │         Includes daily tracking helpers (_seg_agregar_hoja) — no external module
    │
    ▼
Generated Excel files (Archivos_Avance/)
    ├── AVANCE_{YYYY-MM-DD}.xlsx           ← main multi-sheet dashboard
    ├── AVANCE_RESUMIDO.xlsx               ← compact summary for Google Sheets
    └── SEGUIMIENTO_VDD_FIJA_{date}.xlsx   ← VDD1 + VDD2 + VDD3 daily snapshot
    │
    ▼
main_v2.py  —  Distribution orchestrator
    │
    ├── BacksProcess     → Google Sheets (RESUMIDO) + WhatsApp group BACKS
    ├── JefesProcess     → captures TDS!B4:V15 and TDS!Y4:AT17 as images
    │                      + SEGUIMIENTO_VDD_FIJA → WhatsApp group JEFES
    ├── JesusProcess     → AVANCE_{date}.xlsx → WhatsApp contact Jesús
    ├── CristianProcess  → AVANCE_{date}.xlsx → WhatsApp contact Cristian
    ├── GuillermnoProcess→ AVANCE_{date}.xlsx → WhatsApp + email Guillermo
    ├── CarlosProcess    → AVANCE_{date}.xlsx → email Carlos only
    └── ItaloProcess     → MES/DIA sheets → Google Sheets + WhatsApp Italo
            │
            ▼
    wa_client.py (Python HTTP client)
            │
            ▼  HTTP POST (localhost:8002)
    wa_server.js (Node.js + whatsapp-web.js)
            │
            ▼
        WhatsApp Web
```

---

## Modules and responsibilities

### `AVANCE.py` — Main ETL

Core of the system. Takes a period (`YYYY-MM`) and produces all Excel files for the day across 9 clearly marked sections.

Includes daily tracking helpers directly (`_seg_agregar_hoja`), so it has no external module dependency for generating the VDD2 sheet.

**Key functions:**
- `_pedir_periodo()` — prompts for period if not passed as argument
- `_crear_pivot()` — builds pivot tables via xlwings (Excel COM)
- `_color_semaforo()` — applies red (<70%), yellow (70–90%), green (>90%)
- `calc_antiguedad()` — classifies salesperson tenure: `<15d`, `>15d`, `>30d`, `>60d`, `>90d`
- `_normalizar_zonal2()` — derives `zonal2` field: LIMA sub-zones → `"LIMA"`, others copy `zonal`
- `_seg_agregar_hoja()` — generates VDD2 sheet (daily tracking) in any openpyxl workbook

### `main_v2.py` — Orchestrator

Controls the full flow: authenticates Google APIs, waits for the trigger email, runs `AVANCE.py` as a subprocess, then launches the 7 distribution processes in sequence. Marks processed emails to avoid re-triggering.

### `modules/` — Distribution processes

| Module | Channel | What it does |
|--------|---------|--------------|
| `backs_process.py` | WhatsApp | Uploads AVANCE_RESUMIDO to Google Sheets; notifies BACKS group with link |
| `jefes_process.py` | WhatsApp | Captures TDS ranges (`B4:V15` and `Y4:AT17`) as PNG images; sends them with SEGUIMIENTO_VDD_FIJA to JEFES group |
| `jesus_process.py` | WhatsApp | Sends AVANCE_{yesterday}.xlsx to Jesús |
| `cristian_process.py` | WhatsApp | Sends AVANCE_{yesterday}.xlsx to Cristian |
| `guillermo_process.py` | WhatsApp + email | Sends AVANCE_{yesterday}.xlsx via WhatsApp and Gmail (`guillermo_email` from config) |
| `carlos_process.py` | Email | Sends AVANCE_{yesterday}.xlsx via Gmail only (`carlos_email` from config) |
| `italo_process.py` | WhatsApp | Uploads MES and DIA sheets to Google Sheets; sends link to Italo |
| `shared/gmail_helper.py` | — | Sends emails with attachments via Gmail API |
| `msg_utils.py` | — | `pick_variant()`: picks a message variant by date hash to rotate daily texts |

### `generar_resumido.py` — Summary via direct SQL

Generates `AVANCE_RESUMIDO.xlsx` via SQL queries, without opening the main AVANCE file in Excel. The RTCHB/ALTASCHB sheets include CHIMBOTE, NORTE CHICO, and all LIMA* zones.

### `whatsapp_server/wa_server.js` — WhatsApp API

Express server wrapping `whatsapp-web.js` with session persistence via `LocalAuth`. Uses installed Chrome for Puppeteer. Implements a **rate-limited queue** guaranteeing at least 5 seconds between outgoing messages. Mentions are resolved with `getContactById()` with a fallback object for numbers not in the phone's contacts.

### `whatsapp_server/wa_client.py` — Python HTTP client

Python abstraction over the Node.js REST API. Resolves names from `config.json` and retries with exponential backoff (up to 3 attempts).

---

## Project structure

```
AVANCE_MOVISTAR/
│
├── AVANCE.py                     # Main ETL + VDD2 helpers (~3,500 lines)
├── main_v2.py                    # Orchestrator: trigger email → ETL → distribution
├── generar_resumido.py           # Generates AVANCE_RESUMIDO via direct SQL
├── generar_seguimiento_diario.py # Standalone VDD2 script (manual/debug use)
├── run_modulo.py                 # Manual execution of a single module
├── start_wa_server.bat           # Checks/starts wa_server.js each morning
├── stop_wa_server.bat            # Stops wa_server.js
│
├── .env                          # Environment variables (not in git)
├── config.json                   # WhatsApp IDs, messages, paths (not in git)
├── config.example.json           # Configuration template without real data
├── credentials.json              # Google OAuth credentials (not in git)
├── token.json                    # Google OAuth token (not in git)
├── requirements.txt              # Python dependencies
│
├── modules/
│   ├── backs_process.py
│   ├── jefes_process.py
│   ├── jesus_process.py
│   ├── cristian_process.py
│   ├── guillermo_process.py      # Email recipient read from config["guillermo_email"]
│   ├── carlos_process.py         # Email recipient read from config["carlos_email"]
│   ├── italo_process.py
│   ├── msg_utils.py              # Daily message variant rotation
│   └── shared/
│       └── gmail_helper.py       # Email sending via Gmail API
│
├── whatsapp_server/
│   ├── wa_server.js              # Express API + whatsapp-web.js (port 8002)
│   ├── wa_client.py              # Python HTTP client
│   ├── package.json
│   ├── config.json               # Resolved IDs (not in git)
│   └── session_data/             # Persisted WhatsApp session (not in git)
│
└── Archivos_Avance/              # Excel output directory (not in git)
    ├── AVANCE_{date}.xlsx
    ├── AVANCE_RESUMIDO.xlsx
    └── SEGUIMIENTO_VDD_FIJA_{date}.xlsx
```

---

## Installation

### System requirements

- **Windows 10/11** (required for xlwings COM and Task Scheduler)
- **Microsoft Excel** installed
- **Python 3.10+**
- **Node.js 18+**
- **Google Chrome** at `C:\Program Files\Google\Chrome\Application\chrome.exe`
- **ODBC Driver 17 for SQL Server**
- Network access to the corporate SQL Server
- Google account with Gmail, Sheets, and Drive APIs enabled

### Step 1 — Clone the repository

```bash
git clone https://github.com/BlackZero2x/avance_ventas_telcom.git
cd avance_ventas_telcom
```

### Step 2 — Create Python virtual environment

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### Step 3 — Install Node.js dependencies

```bash
cd whatsapp_server
npm install
cd ..
```

### Step 4 — Configure Google credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a project and enable: Gmail API, Google Sheets API, Google Drive API
3. Create OAuth 2.0 credentials → download as `credentials.json`
4. Place `credentials.json` in the project root
5. On first run, a browser window opens for authorization and generates `token.json`

### Step 5 — Create configuration files

```bash
copy config.example.json config.json
```

Create a `.env` file in the project root (see [Configuration](#configuration) for required variables).

### Step 6 — Start the WhatsApp server for the first time

```bash
node whatsapp_server/wa_server.js
```

Scan the QR code with your phone (same as WhatsApp Web). The session is persisted in `whatsapp_server/session_data/`.

### Step 7 — Verify everything works

```bash
python whatsapp_server/wa_client.py --health
python whatsapp_server/wa_client.py --list-groups
python AVANCE.py
```

### Step 8 — Configure Windows Task Scheduler

**Task 1 — WhatsApp server startup**
- Trigger: every business day at 8:20 AM
- Action: run `start_wa_server.bat`

**Task 2 — Main pipeline**
- Trigger: every business day at 8:30 AM
- Action: `python C:\proyectos\AVANCE_MOVISTAR\main_v2.py`
- Enable "Run whether user is logged on or not"
- Enable "Run with highest privileges"

---

## Configuration

### Environment variables (.env)

| Variable | Description |
|----------|-------------|
| `SQL_SERVER` | SQL Server instance (e.g. `SERVERNAME\INSTANCE`) |
| `SQL_DATABASE` | Database name |
| `SQL_USER` / `SQL_PASSWORD` | SQL credentials |
| `MF_CSV_PATH` | Path to local MiFibra CSV file |
| `MF_SHEET_ID` | Private Google Sheet ID for MiFibra |
| `URL_VENTORY` / `URL_RH` / `URL_LCF` | Public Google Sheets CSV export URLs |
| `HTTP_PROXY` / `HTTPS_PROXY` / `NO_PROXY` | Corporate proxy settings (if needed) |
| `AVANCE_DIR` | Project root path |
| `CHECK_INTERVAL_MINUTES` | Orchestrator polling interval (default 10) |

### config.json (based on config.example.json)

```json
{
  "archivos_avance_dir": "C:/proyectos/AVANCE_MOVISTAR/Archivos_Avance",
  "carlos_email": "recipient@yourcompany.com",
  "guillermo_email": "recipient2@yourcompany.com",
  "jefes_tds_rango1": "B4:V15",
  "jefes_tds_rango2": "Y4:AT17",
  "groups": {
    "Group Name": "XXXXXXXX-XXXXXXXX@g.us"
  },
  "contacts": {
    "Contact Name": "51XXXXXXXXX@c.us"
  }
}
```

---

## Manual execution

```bash
# Run the full ETL (prompts for period YYYY-MM)
python AVANCE.py

# Run a single distribution module
python run_modulo.py carlos
python run_modulo.py jefes
python run_modulo.py backs

# WhatsApp client commands
python whatsapp_server/wa_client.py --health
python whatsapp_server/wa_client.py --test
python whatsapp_server/wa_client.py --list-groups
python whatsapp_server/wa_client.py --list-contacts "Name"
python whatsapp_server/wa_client.py --test-image "C:\path\image.png"
python whatsapp_server/wa_client.py --test-file  "C:\path\report.xlsx"
```

### Programmatic usage from Python

```python
import sys
sys.path.insert(0, r"C:\proyectos\AVANCE_MOVISTAR\whatsapp_server")
from wa_client import WhatsAppClient
from modules.msg_utils import pick_variant

wa = WhatsAppClient()
wa.send_text("Group Name", "Message text")
wa.send_file("Group Name", r"C:\path\report.xlsx", caption="Daily report")
wa.send_image("Contact Name", r"C:\path\screenshot.png", caption="TDS dashboard")
wa.send_link("Contact Name", "https://drive.google.com/...", "View in Drive")
wa.send_mention("Group Name", "Hello @number", ["51XXXXXXXXX@c.us"])
```

---

## ETL sections (AVANCE.py)

| # | Section | Description |
|---|---------|-------------|
| 1 | **Load sources** | SQL Server (activations, RT, DITO consultations, previous month) + HR/LCF/quota Excel files + MiFibra CSV + private MiFibra Google Sheet |
| 2 | **Clean sources** | Deduplication, type conversion, text normalization, null handling |
| 3 | **Double join with AppVentory** | Match by `codigo_fe`; fallback by `numero_peticion`; keeps most recent record |
| 4 | **Join with HR** | Adds salesperson name, supervisor, and zone by DNI |
| 5 | **Calculated columns** | `DNI_VENDEDOR` (COALESCE across sources), TV classification, `MATCH_DIRECC`, tenure, week number |
| 6 | **Final renames** | Standardizes column names to the output schema |
| 7 | **LCF join → RIESG** | Adds credit risk score per salesperson |
| 8 | **Final cleanup** | Final deduplication, mandatory field validation |
| 9 | **Excel workbook generation** | TDS (tables 1 & 2, traffic lights), VDD1/VDD2/VDD3, MOVISTAR (daily pivots by `zonal2`), MiFibra, RT/ALTAS sheets with `zonal2`, SEGUIMIENTO_VDD_FIJA |

Only records with `categoria_producto = 'ALTA'` and `fecha_alta IS NOT NULL` are processed.

---

## System outputs

### AVANCE_{YYYY-MM-DD}.xlsx — Main dashboard

| Sheet | Content |
|-------|---------|
| **MOVISTAR** | Daily pivots: conversion, activations (total/regular/flex), speeds, and VDD by `zonal2` |
| **MiFibra** | MiFibra sales and installations dashboard by branch and plan |
| **TDS** | KPIs by zone (table 1, cols B–V) and by supervisor (table 2, cols Y–AT), traffic-light coloring |
| **VDD1** | Per-salesperson detail with Excel formulas |
| **VDD2** | Daily tracking: RT, activations, and consultations per day + last-3-day summary + ALERTS column |
| **VDD3** | Native pivot tables |
| **RT** | Raw total records with `zonal2` field |
| **ALTAS** | Raw activations with `zonal2` field |
| **RH**, **VENTORY**, **CON** | Auxiliary source data |
| **MES**, **DIA** | Data for Italo (hidden) |

### SEGUIMIENTO_VDD_FIJA_{DD-MM-YYYY}.xlsx

Copy of VDD1, VDD2, and VDD3 for the day. Sent to the managers group.

### AVANCE_RESUMIDO.xlsx

Compact version generated via direct SQL (no Excel required). Uploaded to Google Sheets.

---

## WhatsApp server and client

### Available endpoints (wa_server.js, port 8002)

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/health` | GET | Server status (`ready` / `not_ready`) |
| `/list-groups` | GET | Lists all groups with name, ID, and participant count |
| `/list-contacts?name=` | GET | Searches contacts by name |
| `/send-text` | POST | Sends text message (queued, ≥5 s between messages) |
| `/send-image` | POST | Sends image from local path (queued) |
| `/send-file` | POST | Sends file from local path (queued) |
| `/send-mention` | POST | Sends text with mentions; resolves contacts via `getContactById()` with fallback for non-saved numbers |
| `/send-link` | POST | Sends text + URL (queued) |

The `to` field accepts a key name from `config.json` or a direct WhatsApp ID (`51XXXXXXXXX@c.us` for individuals, `XXXXXXXX-XXXXXXXX@g.us` for groups).

---

## WhatsApp ban-risk mitigation

### 1. Rate-limited queue (wa_server.js)
FIFO queue with a minimum **5-second gap** between outgoing messages, regardless of how many Python modules call the server simultaneously.

### 2. Daily message rotation (msg_utils.py)
`pick_variant()` selects a different message variant each day using a date hash. With 3–4 variants per recipient, the text changes daily without manual intervention.

```python
from modules.msg_utils import pick_variant

message = pick_variant(
    variants=[
        "Good morning, please find the daily report attached.",
        "Hi, here is the updated sales report.",
        "Daily report attached — please review.",
    ],
    fallback="Good morning, daily report attached."
)
```

### 3. Extra delays in JefesProcess
- 12 seconds between each TDS screenshot
- 10 seconds before sending the SEGUIMIENTO_VDD_FIJA file

---

## Tech stack

| Layer | Technology | Purpose |
|-------|-----------|---------|
| ETL / data processing | Python 3.10, pandas 2.x | Multi-source joins, transformations, pivots |
| Excel (read/write) | openpyxl | Create and modify .xlsx files, generate VDD2 sheet |
| Advanced Excel automation | xlwings (COM) | Pivot tables, traffic lights, screenshot capture |
| Screenshot capture | Pillow, win32gui | Excel ranges → PNG images |
| Database | SQL Server + pyodbc + SQLAlchemy | Primary activations data source |
| WhatsApp server | Node.js 18 + whatsapp-web.js + Express | Local REST API over WhatsApp Web |
| WhatsApp client | Python + requests | HTTP calls to Node server with retries |
| Google APIs | google-api-python-client | Gmail (trigger), Sheets (data upload), Drive |
| Email delivery | Gmail API | Outbound email to recipients |
| Task scheduling | Windows Task Scheduler | Automatic daily pipeline trigger |

---

## Domain glossary

| Term | Meaning |
|------|---------|
| **ALTA** | A customer who completed the fixed-line service activation process |
| **ALTAS.MF** | MiFibra installations for the period (local CSV + private Google Sheet, deduplicated by DNI + date) |
| **RT** | Registros Totales — all records, not only activations |
| **CON** | DITO consultations (purchase intentions) |
| **TDS** | Technical Data Summary — main dashboard sheet with aggregated KPIs |
| **VDD** | Detalle por Vendedor — individual salesperson performance breakdown |
| **VDD2** | Daily tracking: RT, activations, and CON per day + automatic alerts |
| **zonal2** | Normalized zone field: LIMA sub-zones collapse to `"LIMA"`; others copy `zonal` |
| **RH** | Human Resources — maps each salesperson (DNI) to supervisor and zone |
| **LCF** | Credit risk data source — provides the `RIESG` field |
| **AVANCE** | Quota fulfillment percentage (activations / quota × 100) |
| **eAuren** | SQL Server database where activations are recorded |
| **FE** | Salesperson internal code in AppVentory (join key) |
| **AppVentory** | Corporate sales registration system |
| **BACKS** | Analyst group that receives the summary in Google Sheets |
| **Periodo** | Analysis month in `YYYY-MM` format |
| **D-1** | Morning data corresponds to the previous day |

---

## Troubleshooting

### WhatsApp server won't start
1. Check Node.js: `node --version`
2. Install dependencies: `cd whatsapp_server && npm install`
3. If the session expired, delete `whatsapp_server/session_data/` and re-scan the QR
4. Check `whatsapp_server/logs/wa_server_YYYY-MM-DD.log`

### Mentions appear as plain text (`@51XXXXXXXXX`)
`whatsapp-web.js` requires `Contact` objects in `mentions`, not plain strings. The `/send-mention` endpoint resolves each ID via `getContactById()` with a minimal fallback object for non-saved numbers. If it persists, restart `wa_server.js` to reload the contact list.

### `CopyPicture failed` — TDS screenshots not generated
Occurs when Excel is not the active window. `jefes_process.py` uses `AllowSetForegroundWindow` + `SetForegroundWindow` to mitigate this. If it persists, temporarily set the Task Scheduler task to **"Run only when user is logged on"** to diagnose.

### LIMA shows zero in MOVISTAR sheet
The MOVISTAR pivots group by `zonal2`. Verify that `altas_df` and `rt_df` have the `zonal` column populated before `_normalizar_zonal2()` runs.

### ETL produces Excel but WhatsApp delivery fails
1. Check server: `python whatsapp_server/wa_client.py --health`
2. If it does not respond, run `start_wa_server.bat` manually
3. Review orchestrator logs

### MiFibra Google Sheet data not loading
The process continues with the local CSV only if the Sheet is unavailable (expired token, no network). Verify `token.json` is valid — it refreshes automatically if the refresh token is still active.

---

## License

Internal use. Code is original authorship. Production data (`.env`, `config.json`, `session_data/`, `Archivos_Avance/`, `credentials.json`, `token.json`) is not included in this repository.
