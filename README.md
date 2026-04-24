# Avance Movistar — Sistema de Seguimiento de Ventas

Pipeline de automatización de reportes de ventas para la línea fija de Movistar Perú. Consolida datos de SQL Server y Excel, genera dashboards diarios y distribuye reportes automáticamente por WhatsApp y Google Sheets.

---

## Características principales

- **ETL completo**: extrae datos de SQL Server (`eAuren`), los cruza con fuentes Excel (RH, LCF, cuotas) y produce reportes listos para usar
- **Dashboard Excel multi-hoja**: genera `AVANCE_{fecha}.xlsx` con hojas TDS (KPIs con semáforo), VDD por vendedor, desglose MOVISTAR/MiFibra y seguimiento
- **Distribución automática por WhatsApp**: envía archivos, capturas de pantalla y mensajes a grupos y contactos individuales via API open-wa
- **Sincronización con Google Sheets**: sube datos consolidados para acceso en tiempo real del equipo
- **Orquestación con trigger por email**: detecta un correo específico (Gmail API) y ejecuta todo el pipeline sin intervención manual
- **Anti-ban de WhatsApp**: cola con rate limiting (≥5 s entre envíos), rotación diaria de mensajes (`msg_utils.pick_variant`)

---

## Arquitectura

```
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
        ├── GuillermnoProcess → archivo → contacto WhatsApp
        ├── CarlosProcess   → archivo → contacto WhatsApp
        └── ItaloProcess    → Google Sheets + contacto WhatsApp
                │
                ▼
        wa_client.py → wa_server.js (puerto 8002) → WhatsApp
```

---

## Flujo de ejecución

1. El Programador de Tareas de Windows lanza `main_v2.py` a las 8:30 AM (lunes–viernes)
2. El orquestador autentica Google APIs y verifica el servidor WhatsApp
3. Revisa el correo cada 10 minutos en busca del email trigger
4. Al detectar el trigger: ejecuta `AVANCE.py` para generar los archivos del día
5. Ejecuta los 7 procesos de distribución en cadena
6. Registra el resumen en logs y termina

---

## Stack tecnológico

| Componente | Tecnología |
|-----------|------------|
| ETL / procesamiento | Python 3.10, pandas, openpyxl, xlwings |
| Base de datos | SQL Server 2019, pyodbc, SQLAlchemy |
| Servidor WhatsApp | Node.js, @open-wa/wa-automate |
| Cliente WhatsApp | Python, requests |
| Google APIs | Gmail, Sheets, Drive (google-api-python-client) |
| Automatización Excel | xlwings (COM), openpyxl |
| Captura de pantalla | Pillow, pyautogui |
| Orquestación | Windows Task Scheduler + Python |

---

## Estructura del proyecto

```
AVANCE_MOVISTAR/
├── AVANCE.py                  # ETL principal (9 secciones)
├── main_v2.py                 # Orquestador
├── generar_resumido.py        # Genera AVANCE_RESUMIDO via SQL directo
├── run_modulo.py              # Ejecución manual de módulos individuales
├── start_wa_server.bat        # Inicio y verificación del servidor WhatsApp
├── config.example.json        # Plantilla de configuración (sin datos reales)
├── requirements.txt           # Dependencias Python
├── modules/
│   ├── backs_process.py
│   ├── jefes_process.py
│   ├── jesus_process.py
│   ├── cristian_process.py
│   ├── guillermo_process.py
│   ├── carlos_process.py
│   ├── italo_process.py
│   └── msg_utils.py           # Rotación de mensajes anti-ban
└── whatsapp_server/
    ├── wa_server.js            # API Express (puerto 8002)
    ├── wa_client.py            # Cliente HTTP Python
    └── package.json
```

---

## Instalación y configuración

### Prerrequisitos

- Python 3.10+
- Node.js 18+
- ODBC Driver 17 for SQL Server
- Acceso a SQL Server `AUREN22\AUREN` (base de datos `eAuren`)
- Cuenta de Google con APIs habilitadas (Gmail, Sheets, Drive)
- WhatsApp activo en web.whatsapp.com

### Pasos

```bash
# 1. Clonar el repositorio
git clone https://github.com/TU_USUARIO/avance-movistar.git
cd avance-movistar

# 2. Instalar dependencias Python
pip install -r requirements.txt

# 3. Instalar dependencias Node.js
cd whatsapp_server
npm install

# 4. Configurar credenciales
# - Descargar credentials.json desde Google Cloud Console
# - Copiar config.example.json → config.json y completar con tus datos

# 5. Iniciar servidor WhatsApp (escanear QR la primera vez)
node whatsapp_server/wa_server.js

# 6. Ejecutar el pipeline ETL
python AVANCE.py
```

---

## Secciones del ETL (AVANCE.py)

| Sección | Descripción |
|---------|-------------|
| 1 | Carga desde SQL Server + archivos Excel |
| 2 | Limpieza y normalización de datos |
| 3 | Join doble con AppVentory (por código FE → por petición) |
| 4 | Enriquecimiento con datos de RH (vendedor/supervisor/zona) |
| 5 | Columnas calculadas: antigüedad, semana, clasificaciones |
| 6 | Renombrado de columnas a nombres estándar |
| 7 | Join de score de riesgo LCF |
| 8 | Deduplicación y validación final |
| 9 | Generación del libro Excel: semáforo, tablas dinámicas, VDD |

---

## Mitigación de riesgo de ban en WhatsApp

- **Cola con rate limiting**: mínimo 5 segundos entre cualquier envío saliente
- **Rotación de mensajes**: `msg_utils.pick_variant()` elige una variante distinta cada día usando hash de fecha
- **Delays adicionales**: procesos con alto volumen de medios esperan 10–12 s entre envíos

---

## Glosario

| Término | Significado |
|---------|-------------|
| ALTAS | Clientes activados (ventas completadas) |
| TDS | Resumen Técnico de Datos — hoja principal del dashboard |
| VDD | Detalle por Vendedor |
| RH | Recursos Humanos (mapeo vendedor/supervisor/zona) |
| LCF | Datos de riesgo crediticio |
| AVANCE | Porcentaje de cumplimiento vs cuota |
| eAuren | Base de datos en SQL Server |
| D-1 | Dato del día anterior |

---

## Licencia

Uso interno. El código es de autoría propia; los datos de producción no están incluidos en este repositorio.
