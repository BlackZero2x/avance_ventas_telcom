# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué hace este proyecto

**AVANCE MOVISTAR** es un sistema mensual de seguimiento de ventas para Movistar (Perú). Realiza:
1. Extrae datos de SQL Server + archivos Excel y genera dashboards Excel de múltiples hojas
2. Envía archivos, capturas de pantalla y enlaces de Google Drive a grupos y contactos via WhatsApp

## Patrón para proyectos nuevos en este repo

Cada nueva funcionalidad sigue esta estructura de múltiples pasos:
1. **Script Python principal** — lógica ETL o procesamiento de datos
2. **Cliente Python de WhatsApp** — usa `WhatsAppClient` de `whatsapp_server/wa_client.py` para notificar resultados
3. **Configuración en `config.json`** — agregar grupos/contactos/mensajes nuevos si aplica; usar `msg_utils.pick_variant()` para rotación de mensajes (evitar bans por patrón)
4. Al agregar endpoints nuevos al servidor Node.js, reiniciar `wa_server.js` antes de probar

## Ejecutar el pipeline ETL principal

```bash
# Desde C:\AVANCE_MOVISTAR — solicita el periodo (YYYY-MM) y genera los reportes Excel
python AVANCE.py
```

**Requisitos previos:**
- SQL Server `AUREN22\AUREN`, base de datos `eAuren`, ODBC Driver 17
- Archivos Excel en el directorio de trabajo: `rh.xlsx`, `lcf.xlsx` (opcional: `cuotas.xlsx`)
- Paquetes Python (ver `requirements.txt`): `pandas`, `sqlalchemy`, `pyodbc`, `openpyxl`, `xlwings`, `google-api-python-client`, `pywin32`

## Servidor y cliente de WhatsApp

```bash
# Iniciar servidor (puerto 8002) — escanear QR solo la primera vez
cd whatsapp_server
npm install
node wa_server.js

# Verificar que está listo
python whatsapp_server/wa_client.py --health
```

**Comandos de prueba del cliente:**
```bash
python whatsapp_server/wa_client.py --test                        # texto a tu propio número
python whatsapp_server/wa_client.py --list-groups                 # listar grupos
python whatsapp_server/wa_client.py --list-contacts "Nombre"      # buscar contacto
python whatsapp_server/wa_client.py --test-image "C:\ruta\img.png"
python whatsapp_server/wa_client.py --test-file  "C:\ruta\archivo.xlsx"
python whatsapp_server/wa_client.py --test-link  "https://..."
```

**Uso programático desde otro script Python:**
```python
import sys
sys.path.insert(0, r"C:\AVANCE_MOVISTAR\whatsapp_server")
from wa_client import WhatsAppClient
from msg_utils import pick_variant

wa = WhatsAppClient()
wa.send_text("Canal Fija 2026", "Mensaje")
wa.send_file("Back de AUREN 2025", r"C:\ruta\reporte.xlsx", caption="Reporte actualizado")
wa.send_image("Cristian", r"C:\ruta\captura.png", caption="TDS del día")
wa.send_link("Jesús", "https://drive.google.com/...", "Seguimiento FIJA")
```

El destinatario (`to`) puede ser un nombre de `config.json` o un ID directo de WhatsApp (ej. `51962969371@c.us`).

## Arquitectura

### Flujo de datos

```
SQL Server (eAuren) + Excel (rh.xlsx, lcf.xlsx, cuotas.xlsx)
        ↓
    AVANCE.py  (ETL: ~3.500 líneas, 9 secciones numeradas)
        ↓
    Salida Excel: AVANCE_RESUMIDO.xlsx + SEGUIMIENTO_VDD_FIJA_DD-MM-YYYY.xlsx
        ↓
    wa_client.py  →  wa_server.js (puerto 8002)  →  WhatsApp
```

### AVANCE.py — Secciones de procesamiento

| Sección | Propósito |
|---------|-----------|
| 1 | Carga desde SQL (`fija_registros_totales`, `fija_altas`, etc.) + fuentes Excel |
| 2 | Limpieza/normalización de nulos y campos |
| 3 | Join de datos de ventory por código FE o número de petición |
| 4 | Enriquecimiento con RH (vendedor/supervisor/zona) |
| 5 | Cálculo de antigüedad (`<15d`, `>15d`, `>30d`, `>60d`, `>90d`), semana, normalizaciones |
| 6 | Renombrado de columnas a nombres estándar finales |
| 7 | Join de score de riesgo LCF (`RIESG`) |
| 8 | Deduplicación y validación final |
| 9 | Generación del libro Excel con múltiples hojas, tablas dinámicas y semáforo de colores |

**Funciones auxiliares clave en AVANCE.py:**
- `_pedir_periodo()` — solicita al usuario el periodo YYYY-MM
- `_crear_pivot()` — construye tablas dinámicas via xlwings
- `_color_semaforo()` — aplica colores rojo (<70%) / amarillo (70–90%) / verde (>90%)
- `calc_antiguedad()` — clasifica la antigüedad del vendedor

### Estructura de salida Excel

- **MOVISTAR / MiFibra** — datos crudos separados por servicio
- **TDS** — dashboard principal de KPIs: rendimiento vs cuota, métricas por supervisor, semáforo, tablas dinámicas
- **VDD1 / VDD2 / VDD3** — desglose detallado por vendedor con fórmulas Excel
- **SEGUIMIENTO_VDD_FIJA** — copia snapshot de las hojas VDD para seguimiento con interesados

### Componentes de WhatsApp

- `wa_server.js` — API Express (puerto 8002) sobre `@open-wa/wa-automate`; persiste sesión en `session_data/`; logs en `logs/wa_server_YYYY-MM-DD.log`. **Implementa cola con rate limiting: mínimo 5 segundos entre envíos cualquiera (evita bans por ráfagas rápidas)**
- `wa_client.py` — cliente HTTP Python; resuelve nombres desde `config.json`; reintenta con backoff exponencial (3 intentos)
- `msg_utils.py` — función `pick_variant(variants, fallback)` que elige una variante de mensaje por día (hash de fecha) para evitar patrones repetitivos que activen spam-detectors de WhatsApp
- `config.json` — IDs de WhatsApp para grupos y contactos, más plantillas de mensajes; incluye `*_variants` arrays con múltiples opciones de texto para cada destinatario

**Endpoints disponibles en wa_server.js:**

| Endpoint | Método | Descripción |
|----------|--------|-------------|
| `/health` | GET | Estado del servidor |
| `/list-groups` | GET | Lista todos los grupos |
| `/list-contacts?name=` | GET | Busca contactos por nombre |
| `/send-text` | POST | Envía texto (encolado con delay) |
| `/send-image` | POST | Envía imagen desde ruta local (base64, encolado) |
| `/send-file` | POST | Envía archivo (xlsx, pdf, csv…) desde ruta local (encolado) |
| `/send-mention` | POST | Envía texto con menciones (encolado) |
| `/send-link` | POST | Envía texto + URL (encolado) |

**Mitigación de riesgo de ban:** El servidor Node.js implementa una cola que garantiza ≥5 s entre cualquier envío saliente, sin importar la cantidad de procesos que llamen en paralelo. Los módulos Python también insertan delays adicionales: `jefes_process.py` espera 12 s entre imágenes y 10 s antes del archivo (grupo con mayor volumen mediático). Todos los módulos usan `pick_variant()` para rotar textos diarios y evitar patrones detectables. Ver https://github.com/rmyndharis/OpenWA/blob/main/docs/16-risk-management.md para contexto de seguridad.

**Nota Windows:** La consola usa cp1252. Los prints con emojis dan `UnicodeEncodeError` — usar `[OK]`/`[ERROR]` en lugar de emojis en mensajes de consola. Para mostrar nombres de grupos con caracteres especiales usar `.encode('cp1252', errors='replace').decode('cp1252')`.

## Flujo de ejecución

El orquestador `main_v2.py` se dispara vía Programador de Tareas de Windows (lunes-viernes, 8:30 AM):

1. **Autentica Google APIs** (Gmail, Sheets, Drive)
2. **Revisa cada 10 min si llegó el email trigger** desde `e@auren.com.pe` con asunto `"avance_ventas - Actualización disponible"`
3. Al detectar trigger: ejecuta `AVANCE.py` para generar archivos del día
4. Ejecuta 6 procesos en cadena (cada uno espera 3 s tras completar):
   - `BacksProcess`: genera AVANCE_RESUMIDO → sube a Google Sheets → notifica grupo BACKS
   - `JefesProcess`: captura rangos TDS de Excel → envía imágenes + archivo SEGUIMIENTO_VDD a grupo JEFES (con delays: 12 s entre imágenes, 10 s antes de archivo)
   - `JesusProcess`: envía AVANCE_{fecha}.xlsx a Jesús
   - `CristianProcess`: envía AVANCE_{fecha}.xlsx a Cristian
   - `GuillermnoProcess`: envía AVANCE_{fecha}.xlsx a Guillermo
   - `ItaloProcess`: sube hojas MES/DIA a Google Sheets → notifica a Italo
5. Marca emails procesados para evitar duplicados
6. Termina (no persiste indefinidamente)

**Nota:** La verificación de WhatsApp cada mañana (`start_wa_server.bat` ejecutado por Programador de Tareas) envía notificación WhatsApp confirmando el servidor está activo.

## Estructura de módulos

- `modules/` — scripts de procesos (backs, jefes, jesus, cristian, guillermo, italo) + `msg_utils.py` (rotación de mensajes)
- `whatsapp_server/` — servidor Node.js (`wa_server.js`) + cliente Python (`wa_client.py`) + config
- `AVANCE.py` — ETL principal (~3.5k líneas, 9 secciones numeradas)
- `generar_resumido.py` — genera AVANCE_RESUMIDO via SQL directo (sin Excel COM)
- `main_v2.py` — orquestador: autentica Google, detecta trigger, ejecuta procesos

## Glosario de términos del dominio

| Término | Significado |
|---------|-------------|
| ALTAS | Clientes activados (ventas completadas) |
| TDS | Resumen Técnico de Datos — hoja principal del dashboard |
| VDD | Detalle por Vendedor — desglose de rendimiento individual |
| RH | Recursos Humanos (mapeo vendedor/supervisor/zona) |
| LCF | Datos de riesgo crediticio (campo `RIESG`) |
| AVANCE | Porcentaje de cumplimiento vs cuota |
| eAuren | Nombre de la base de datos en SQL Server |
| FE | Código interno del vendedor usado para los joins |
| AppVentory | Sistema de registro de altas; fuente de datos sobre vendedor/petición |
