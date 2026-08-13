# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Cambios en wa_server.js (18/07/2026 — estabilidad de la cola global)

El servidor `whatsapp_server/wa_server.js` (puerto 8002) es **compartido por 5 proyectos**:
AVANCE_MOVISTAR, SSFF (cortes + pedidos), Dash_consultas_pbi, VPN_MIFIBRA (captura de corte) y tablas Jesús.
Tras las fallas del 15–17/07 (crash-loop de Chrome + "cola llena" por horas) se aplicó:

- **Cola priorizada y drenable** (reemplaza la cadena de promesas anterior): array explícito con
  prioridad. Campo opcional `priority` en el body de los endpoints de envío
  (`"avance"` o número ≥10 = alta prioridad; default normal). AVANCE, al ser un evento
  puntual por trigger de correo, debe enviar con `priority:"avance"` para pasar delante de
  los cortes horarios de SSFF. **Al detectar contexto muerto (Promise collected / Protocol
  error / detached Frame) la cola se DRENA de golpe** (rechaza todos los pendientes al
  instante) en vez de esperar 5 s por cada uno — esto elimina la saturación de "cola llena".
- **Chrome multiproceso**: se quitaron `--single-process` y `--no-zygote` de los args de
  puppeteer (eran la causa raíz de "Promise was collected" bajo carga sostenida).
- **Reinicio preventivo cada 6 h** (antes 1x nocturno), solo con cola vacía.
- **Instancia única**: si el puerto 8002 ya está en uso, la nueva instancia aborta con
  `EADDRINUSE` (evita dos clientes WA compitiendo por la misma sesión). El
  `watchdog_wa_server.bat` mata node huérfano antes de relanzar.

> Reiniciar `wa_server.js` como administrador para aplicar estos cambios (el proceso lo lanza
> el Task Scheduler con privilegios elevados). Backup del original: `wa_server.js.bak_20260718`.

### Integración VPN_MIFIBRA (captura de corte por WhatsApp)

`C:\proyectos\VPN_MIFIBRA\enviar_captura_vpn.bat` genera el PNG del corte
(`pivot_corte_dia.py --sin-descarga`) y lo envía con `enviar_captura_vpn.py` +
`wa_client.py` al servidor 8002. Destino y textos en `VPN_MIFIBRA/config.json`
(clave `captura_corte`). **Destino actual: número propio de prueba** (`51975155264@c.us`).
Tareas programadas `VPN_Captura_*` (10:15, 12:40, 14:40, 16:40, 18:40, 20:40), 10 min
después de cada sync VPN. Prioridad `normal` (no compite con el trigger de AVANCE).

  Antes de escribir o modificar cualquier código de este proyecto, aplica
  siempre estas reglas de confidencialidad:

  1. CREDENCIALES: Nunca escribas valores reales (contraseñas, usuarios,
     servidores, tokens, API keys) directamente en el código. Toda credencial
     debe leerse desde variables de entorno o un archivo .env, sin valores
     por defecto que revelen datos reales. Si falta una variable requerida,
     el script debe abortar con un mensaje de error claro.

  2. DATOS PERSONALES: No incluyas nombres reales de personas, correos,
     teléfonos ni códigos de clientes en el código fuente. Usa variables
     de entorno o archivos de configuración externos.

  3. INFRAESTRUCTURA: Evita hardcodear nombres de servidores, bases de datos,
     DSNs, rutas de red internas o convenciones de codificación internas
     (como tipos de venta o estados) en lugares visibles del código. Si son
     necesarios para el funcionamiento, centralízalos en un bloque de
     configuración claramente marcado como "ajustar en cada entorno".

  4. ANTES DE VERSIONAR: Cuando vayas a preparar código para Git, revisa
     activamente si hay credenciales, datos de clientes o información de
     infraestructura interna que deba moverse a .env o eliminarse.

## Qué hace este proyecto

**AVANCE MOVISTAR** es un sistema mensual de seguimiento de ventas para Movistar (Perú). Realiza:
1. Extrae datos de SQL Server + archivos Excel/CSV y genera dashboards Excel de múltiples hojas
2. Envía archivos, capturas de pantalla y enlaces de Google Drive a grupos y contactos via WhatsApp
3. Envía archivos por correo electrónico (Gmail API) a destinatarios específicos

## Patrón para proyectos nuevos en este repo

Cada nueva funcionalidad sigue esta estructura de múltiples pasos:
1. **Script Python principal** — lógica ETL o procesamiento de datos
2. **Cliente Python de WhatsApp** — usa `WhatsAppClient` de `whatsapp_server/wa_client.py` para notificar resultados
3. **Configuración en `config.json`** — agregar grupos/contactos/mensajes nuevos si aplica; usar `msg_utils.pick_variant()` para rotación de mensajes (evitar bans por patrón)
4. Al agregar endpoints nuevos al servidor Node.js, reiniciar `wa_server.js` antes de probar

## Ejecutar el pipeline ETL principal

```bash
# Desde C:\proyectos\AVANCE_MOVISTAR — solicita el periodo (YYYY-MM) y genera los reportes Excel
python AVANCE.py
```

**Requisitos previos:**
- SQL Server `AUREN22\AUREN`, base de datos `eAuren`, ODBC Driver 17
- Archivos Excel en el directorio de trabajo: `rh.xlsx` (opcional: `cuotas.xlsx`, `cuotas_zonal_sup.xlsx`)
- Entorno virtual compartido: `C:\proyectos\.venv\` (paquetes en `C:\proyectos\requirements.txt`)
- Credenciales en `.env`: `SQL_SERVER`, `SQL_DATABASE`, `SQL_USER`, `SQL_PASSWORD`
- Tabla SQL MiFibra: `[dbo].[mifibra_ventas_hora]` en la misma base `eAuren` (mismas credenciales SQL)
- Google Sheet MiFibra privado: ID en `.env` como `MF_SHEET_ID`
- Google credentials: `credentials.json` + `token.json` en la raíz del proyecto

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
sys.path.insert(0, r"C:\proyectos\AVANCE_MOVISTAR\whatsapp_server")
from wa_client import WhatsAppClient
from msg_utils import pick_variant

wa = WhatsAppClient()
wa.send_text("Canal Fija 2026", "Mensaje")
wa.send_file("Back de AUREN 2025", r"C:\ruta\reporte.xlsx", caption="Reporte actualizado")
wa.send_image("Cristian", r"C:\ruta\captura.png", caption="TDS del día")
wa.send_link("Jesús", "https://drive.google.com/...", "Seguimiento FIJA")
wa.send_mention("Canal Fija 2026", "Mensaje @número", ["51962969371@c.us"])
```

El destinatario (`to`) puede ser un nombre de `config.json` o un ID directo de WhatsApp (ej. `51962969371@c.us`).

## Arquitectura

### Flujo de datos

```
SQL Server (eAuren) + Excel (rh.xlsx, cuotas.xlsx, cuotas_zonal_sup.xlsx, integratel_riesgo.xlsx)
+ CSV local (BD_Ventas_AUREN.csv) + Google Sheet privado (MF_SHEET_ID)
        ↓
    AVANCE.py  (ETL: ~3.500 líneas, 9 secciones numeradas)
        ↓
    Salida Excel: AVANCE_{fecha}.xlsx + AVANCE_RESUMIDO.xlsx + SEGUIMIENTO_VDD_FIJA_DD-MM-YYYY.xlsx
        ↓
    wa_client.py  →  wa_server.js (puerto 8002)  →  WhatsApp
    gmail_helper.py  →  Gmail API  →  Correo electrónico
```

### AVANCE.py — Secciones de procesamiento

| Sección | Propósito |
|---------|-----------|
| 1 | Carga desde SQL (`fija_registros_totales`, `fija_altas`, etc.) + fuentes Excel + CSV MiFibra + Google Sheet MiFibra |
| 2 | Limpieza/normalización de nulos y campos |
| 3 | Join de datos de Ventory por código FE o número de petición |
| 4 | Enriquecimiento con RH (vendedor/supervisor/zona) |
| 5 | Cálculo de antigüedad (`<15d`, `>15d`, `>30d`, `>60d`, `>90d`), semana, normalizaciones |
| 6 | Renombrado de columnas a nombres estándar finales |
| 7 | Join de riesgo Integratel por ORDER_KEY↔peticion (`RIESG`) |
| 8 | Deduplicación y validación final |
| 9 | Generación del libro Excel con múltiples hojas, tablas dinámicas y semáforo de colores |

**Funciones auxiliares clave en AVANCE.py:**
- `_pedir_periodo()` — solicita al usuario el periodo YYYY-MM
- `_crear_pivot()` — construye tablas dinámicas via xlwings
- `_color_semaforo()` — aplica colores rojo (<70%) / amarillo (70–90%) / verde (>90%)
- `calc_antiguedad()` — clasifica la antigüedad del vendedor
- `_normalizar_zonal2(df)` — renombra `DEPARTAMENTO` → `zonal2`; para zonales con prefijo "LIMA" escribe "LIMA", para el resto copia el valor de `zonal`

### Estructura de salida Excel

El archivo principal `AVANCE_{fecha}.xlsx` contiene las hojas en este orden:

| Hoja | Descripción |
|------|-------------|
| **MOVISTAR** | Datos crudos de altas Movistar |
| **MiFibra** | Dashboard MiFibra: ventas e instalaciones por filial y plan |
| **TDS** | Dashboard principal: KPIs por zonal (tabla 1) y por supervisor (tabla 2 y 3), semáforo, tablas dinámicas |
| **VDD1** | Detalle por vendedor con fórmulas Excel |
| **VDD2** | Seguimiento diario por vendedor: RT, ALTAS y CONSULTAS por día + columnas resumen últimos 3 días + ALERTAS |
| **VDD3** | Tablas dinámicas nativas |
| **RT** | Registros totales crudos (con campo `zonal2` normalizado) |
| **ALTAS** | Altas crudas (con campo `zonal2` normalizado) |
| **RH**, **VENTORY**, **CON** | Fuentes auxiliares |
| **MES**, **DIA** | Datos para Italo (ocultas) |

`SEGUIMIENTO_VDD_FIJA_{fecha}.xlsx` contiene: VDD1, VDD2 (seguimiento diario), VDD3.

### Campo zonal2 (RT y ALTAS)

Las hojas RT y ALTAS tienen el campo `zonal2` (renombrado desde `DEPARTAMENTO`):
- Si `zonal` empieza con `"LIMA"` → `zonal2 = "LIMA"`
- Cualquier otra zonal → `zonal2 = zonal` (copia exacta)

Las fórmulas de TDS tabla 1 (`AVANCEMES`, días, `%Conver`, `%FLEX`) usan `Tbl_ALTAS[zonal2]` y `Tbl_RT[zonal2]` como criterio de filtro.

### Hoja VDD2 — Seguimiento diario

Generada por `generar_seguimiento_diario.py` (importable o standalone). Estructura:

- **Fila 1**: bloques de color (DATOS DE VENDEDORES | REGISTROS TOTALES | ALTAS | CONSULTAS | RESUMEN ULTIMOS 3 DIAS | ALERTAS)
- **Fila 2**: nombres de columna
- **Filas 3+**: una fila por vendedor

Columnas de datos VDD (A–G): ZONAL, SUPERVISOR, DNI, VENDEDOR, ANTIG, F_INGRESO (formato DD/MM/YYYY), ESQUEMA

Columnas de métricas: `RT_D01`…`RT_Dnn`, `TOTAL_RT`, `ALT_D01`…`ALT_Dnn`, `TOTAL_ALT`, `CON_D01`…`CON_Dnn`, `TOTAL_CON`

Columnas extra al final:
- `RT_ULT_3D` — suma de los 3 días más recientes de RT
- `CON_ULT_3D` — suma de los 3 días más recientes de CON
- `ALT_ULT_3D` — suma de los 3 días más recientes de ALT
- `ALERTAS` — fondo amarillo, negrita negra:
  ```
  =IF(CON_ULT_3D=0,"SIN CON ULT3D",IF(RT_ULT_3D=0,"SIN RT ULT3D",
   IF(ALT_ULT_3D=0,"SIN ALTAS ULT3D",CONCATENATE(TOTAL_ALT," ALT"))))
  ```

Fuente global: Aptos Narrow 11.

### Fuente de datos MiFibra (ALTAS.MF en VDD1)

Dos fuentes combinadas:
1. **SQL Server** — tabla `[dbo].[mifibra_ventas_hora]` en `eAuren` (mismas credenciales que el pipeline principal). Histórico versionado por contrato (`es_actual=1` = versión vigente de cada `numcontrato`). VENTAS: filtro por `fechainscripcionficha` (renombrado `FECHA DE VENTA`) = mes actual. INSTALADAS: filas con `fechainstinternet` (renombrado `FECHA DE INSTALACION`) no nula. `PLAN FINAL` se mapea desde `paqueteinicialinternet`. Normalización FILIAL: ANCASH→CHIMBOTE, LA LIBERTAD→TRUJILLO.
2. **Google Sheet privado** — `MF_SHEET_ID` en `.env`. Hoja `"MiFibra"`. Actúa como lookup `NUMERO CONTRATO → DNI_vendedor`. Acceso autenticado con `token.json`.

Si SQL falla, `_mf_raw` queda vacío y el pipeline continúa sin datos MiFibra (no aborta).

### Fuente de datos Integratel (RIESG en RT/ALTAS)

Excel enviado por correo desde `eduardo.pinco@integratel.com.pe`, asunto fijo
`"Reporte de casos observados sujetos a PRE-PENALIDAD - Socio AUREN (MASIVO)"`.
`main_v2.py` / `run_modulo.py` descargan el adjunto más reciente (vía
`modules/shared/integratel_helper.py`) **antes** de invocar `AVANCE.py`, y lo
**acumulan** sobre `integratel_riesgo.xlsx` (cada adjunto trae solo casos
recientes, no el histórico completo). Se deduplica por `ORDER_KEY`, ganando la
fila del adjunto nuevo si el mismo `ORDER_KEY` ya estaba acumulado. Así el
archivo sigue sirviendo entre meses — el filtro real por período lo aplica
`AVANCE.py` al comparar contra `peticion` del mes actual, así que las filas
viejas no afectan el resultado. Si no llega correo nuevo, el archivo
acumulado existente se conserva sin cambios.

`AVANCE.py` lee la primera hoja del Excel y usa el campo `ORDER_KEY`: si el
`peticion` de un registro de RT o ALTAS aparece en `ORDER_KEY`, `RIESG = 1`;
si no, `RIESG = 0`. Reemplaza por completo al antiguo join con LCF
(`URL_LCF`, obsoleto). Si `integratel_riesgo.xlsx` no existe, `RIESG` queda
en 0 para todos y se imprime un aviso.

### TDS — Tablas y captura de pantalla

- **Tabla 1** (cols B–V): métricas por ZONAL. HC (col O/P) usa `SUMIF(Y:Y, B{r}, AM:AM)` — criterio exacto igual al nombre de la zonal en tabla 2.
- **Tabla 2** (cols Y–AT): métricas por SUPERVISOR. Datos leídos dinámicamente de `cuotas_zonal_sup.xlsx` hoja `SUPERVISOR`. Bordes: fila penúltima sin borde inferior, última fila (totales) con borde inferior.
- **Captura enviada a jefes**: `jefes_tds_rango1 = "B4:AD15"` y `jefes_tds_rango2 = "AG4:BJ14"` (configurables en `config.json`).

### Componentes de WhatsApp

- `wa_server.js` — API Express (puerto 8002) sobre **`whatsapp-web.js`** (migrado desde @open-wa); persiste sesión en `session_data/` via `LocalAuth`; usa Chrome instalado (`executablePath`); logs en `logs/wa_server_YYYY-MM-DD.log`. **Cola con rate limiting: mínimo 5 s entre envíos.**
- `wa_client.py` — cliente HTTP Python; resuelve nombres desde `config.json`; reintenta con backoff exponencial (3 intentos)
- `msg_utils.py` — función `pick_variant(variants, fallback)` que elige una variante de mensaje por día (hash de fecha)
- `config.json` — IDs de WhatsApp para grupos y contactos, plantillas de mensajes, rangos TDS, IDs de menciones

**Menciones en whatsapp-web.js:** El endpoint `/send-mention` resuelve cada ID con `getContactById()`. Si el contacto no está en la agenda del teléfono, usa un objeto de fallback con la estructura mínima `{ id: { _serialized, user, server } }` — esto garantiza que la mención se renderice aunque el número no esté guardado.

**Endpoints disponibles en wa_server.js:**

| Endpoint | Método | Descripción |
|----------|--------|-------------|
| `/health` | GET | Estado del servidor |
| `/list-groups` | GET | Lista todos los grupos |
| `/list-contacts?name=` | GET | Busca contactos por nombre |
| `/send-text` | POST | Envía texto (encolado con delay) |
| `/send-image` | POST | Envía imagen desde ruta local (encolado) |
| `/send-file` | POST | Envía archivo desde ruta local (encolado) |
| `/send-mention` | POST | Envía texto con menciones (resuelve contactos, encolado) |
| `/send-link` | POST | Envía texto + URL (encolado) |

**Mitigación de riesgo de ban:** Cola Node.js garantiza ≥5 s entre envíos. `jefes_process.py` espera 12 s entre imágenes y 10 s antes del archivo. Todos los módulos usan `pick_variant()` para rotar textos diarios.

**Nota Windows:** La consola usa cp1252. Los prints con emojis dan `UnicodeEncodeError` — usar `[OK]`/`[ERROR]`. Para mostrar nombres con caracteres especiales usar `.encode('cp1252', errors='replace').decode('cp1252')`.

## Flujo de ejecución

El orquestador `main_v2.py` se dispara vía Programador de Tareas de Windows (lunes-viernes, 8:30 AM):

1. **Autentica Google APIs** (Gmail, Sheets, Drive) con `credentials.json` + `token.json`
2. **Revisa cada 10 min si llegó el email trigger** desde `e@auren.com.pe` con asunto `"avance_ventas - Actualización disponible"`
3. Al detectar trigger: ejecuta `AVANCE.py` para generar archivos del día
4. Ejecuta 7 procesos en cadena (cada uno espera 3 s tras completar):
   - `BacksProcess`: genera AVANCE_RESUMIDO → sube a Google Sheets → notifica grupo BACKS
   - `JefesProcess`: captura rangos TDS (`B4:V15` y `Y4:AT17`) → envía imágenes + SEGUIMIENTO_VDD_FIJA al grupo JEFES
   - `JesusProcess`: envía AVANCE_{fecha}.xlsx a Jesús por WhatsApp
   - `CristianProcess`: envía AVANCE_{fecha}.xlsx a Cristian por WhatsApp
   - `GuillermnoProcess`: envía AVANCE_{fecha}.xlsx a Guillermo por WhatsApp **y por correo** a `guillermoj.hinostroza@auren.com.pe`
   - `CarlosProcess`: envía AVANCE_{fecha}.xlsx a Carlos **solo por correo** a `carlos.parra@auren.com.pe`
   - `ItaloProcess`: sube hojas MES/DIA a Google Sheets → notifica a Italo
5. Marca emails procesados para evitar duplicados
6. Termina (no persiste indefinidamente)

**Correo electrónico:** `GmailHelper` en `modules/shared/gmail_helper.py`. Usa Gmail API con la cuenta corporativa `augusto.moreno@auren.com.pe`. Nunca usar `amorenop@outlook.com` (es personal/GitHub únicamente).

**Nota:** `start_wa_server.bat` ejecutado por Programador de Tareas cada mañana envía notificación WhatsApp confirmando que el servidor está activo.

## Estructura de módulos

- `modules/` — scripts de procesos (backs, jefes, jesus, cristian, guillermo, carlos, italo) + `msg_utils.py` + `shared/gmail_helper.py`
- `whatsapp_server/` — servidor Node.js (`wa_server.js`) + cliente Python (`wa_client.py`) + `config.json`
- `AVANCE.py` — ETL principal (~3.5k líneas, 9 secciones numeradas)
- `generar_resumido.py` — genera AVANCE_RESUMIDO via SQL directo; hojas RTCHB/ALTASCHB incluyen zonales CHIMBOTE, NORTE CHICO y todas las que empiezan con "LIMA"
- `generar_seguimiento_diario.py` — genera la hoja VDD2 (seguimiento diario); exportable como `agregar_hoja_seguimiento(wb, avance_path, periodo, engine)`
- `main_v2.py` — orquestador: autentica Google, detecta trigger, ejecuta procesos
- `run_modulo.py` — ejecutor manual de módulos individuales; autentica Google y pasa `gmail_service` a los módulos que lo requieren

## Variables de entorno (.env)

| Variable | Descripción |
|----------|-------------|
| `SQL_SERVER` | Servidor SQL (`AUREN22\AUREN`) |
| `SQL_DATABASE` | Base de datos (`eAuren`) |
| `SQL_USER` / `SQL_PASSWORD` | Credenciales SQL |
| ~~`MF_CSV_PATH`~~ | Obsoleto — reemplazado por `[dbo].[mifibra_ventas_hora]` en SQL |
| `MF_SHEET_ID` | ID del Google Sheet privado de MiFibra |
| `URL_VENTORY` | URL CSV pública de Google Sheets (VENTORY) |
| ~~`URL_LCF`~~ | Obsoleto — RIESG ahora viene de `integratel_riesgo.xlsx` (ORDER_KEY↔peticion) |
| `HTTP_PROXY` / `HTTPS_PROXY` / `NO_PROXY` | Proxy corporativo |
| `AVANCE_DIR` | Ruta raíz del proyecto |
| `CHECK_INTERVAL_MINUTES` | Intervalo de polling del orquestador (default 10) |
| `OWNER_WA_ID` | Número WA del responsable técnico (`51XXXXXXXXX@c.us`). Recibe alertas internas y el send-test periódico cada 3 horas. |

## Glosario de términos del dominio

| Término | Significado |
|---------|-------------|
| ALTAS | Clientes activados (ventas completadas) |
| ALTAS.MF | Instalaciones MiFibra del período (combinación CSV + Google Sheet) |
| TDS | Resumen Técnico de Datos — hoja principal del dashboard |
| VDD | Detalle por Vendedor — desglose de rendimiento individual |
| VDD2 | Seguimiento diario (reemplaza la antigua hoja de pivot diario) |
| zonal2 | Campo normalizado: "LIMA" para cualquier sub-zonal LIMA; igual a `zonal` para el resto |
| RH | Recursos Humanos (mapeo vendedor/supervisor/zona) |
| RIESG | Score de riesgo (1/0), calculado por cruce ORDER_KEY↔peticion con el Excel de Integratel |
| AVANCE | Porcentaje de cumplimiento vs cuota |
| eAuren | Nombre de la base de datos en SQL Server |
| FE | Código interno del vendedor usado para los joins |
| AppVentory | Sistema de registro de altas; fuente de datos sobre vendedor/petición |
| RT | Registros Totales (todos los registros, no solo ALTAS) |
| CON | Consultas DITO (intenciones de compra) |
