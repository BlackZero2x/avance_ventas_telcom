# Asistencia Planilla — CLAUDE.md

Subproyecto independiente dentro de AVANCE MOVISTAR. Genera y envía el informe diario de asistencia de vendedores PLANILLA vía WhatsApp.

## Qué hace

- Lee la hoja `RRHH` del Google Sheet de asistencia vía OAuth (filtra PLANILLA + ACTIVO + EN CAMPO)
- Lee la hoja `ASISTENCIA_BBDD` del mismo Sheet para obtener los registros del día
- Cruza ambas fuentes por DNI y calcula tablas por ZONAL y por SUPERVISOR
- Genera dos imágenes Excel formateadas (via openpyxl + xlwings)
- Envía las imágenes al grupo `Canal Fija 2026 Gestión AUREN` con menciones a Jesús Ascencios y Carlos P

## Cómo ejecutar

```bash
# Informe de hoy al grupo
python asistencia_planilla/asistencia_planilla.py

# Fecha específica (para pruebas)
python asistencia_planilla/asistencia_planilla.py --fecha 2026-05-27

# Prueba a número propio en lugar del grupo
python asistencia_planilla/asistencia_planilla.py --destino "51975155264@c.us"

# Combinado
python asistencia_planilla/asistencia_planilla.py --fecha 2026-05-27 --destino "51975155264@c.us"
```

## Rutas importantes

El script vive en `asistencia_planilla/asistencia_planilla.py` pero usa recursos de la raíz del proyecto:
- `BASE_DIR = Path(__file__).parent.parent` — apunta a la raíz
- `config.json` → raíz del proyecto
- `token.json`, `credentials.json` → raíz
- `logs/` → raíz (`asistencia_YYYYMMDD.log`)
- `temp/` → raíz (PNGs temporales)

## Google Sheet de asistencia

- **Sheet ID**: `1EGNnQYG51MROVf2tuxKmEqCkw9lfIXQqJBkJke3AITk`
- **Hoja RRHH** (gid 241856834): padrón de vendedores. Columnas clave: `DNI`, `ESQUEMA`, `ESTADO`, `FEEDBACK_RH`, `ZONA`, `SUPERVISOR`
- **Hoja ASISTENCIA_BBDD** (gid 1135922988): registros diarios de asistencia. Columnas clave: `DNI`, `DAY` (fecha en formato DD/MM/YYYY)

## Filtros aplicados sobre RRHH

1. `ESQUEMA == PLANILLA` (PART-TIME se mapea a PLANILLA)
2. `ESTADO == ACTIVO`
3. `FEEDBACK_RH == EN CAMPO`

## Paleta de colores

| Elemento | Color | Hex |
|---|---|---|
| Encabezados y fila TOTAL | Azul cielo | `#118AB2` |
| Filas alternas | Azul cielo muy claro | `#EBF7FB` |
| Semáforo ≥ 90% | Verde pastel | `#83EBB0` |
| Semáforo ≥ 70% | Amarillo pastel | `#FFE87F` |
| Semáforo < 70% | Rojo pastel | `#F7A3B7` |

El texto del campo `%ASISTENCIA` es siempre negro en negrita.

## Menciones

Los IDs de WhatsApp para mencionar se leen desde `config.json` bajo la clave `asistencia_menciones`:
```json
"asistencia_menciones": ["51944956042@c.us", "51968035020@c.us"]
```
Corresponden a Jesús Ascencios y Carlos P. Se envían en un mensaje de texto separado tras la primera imagen.

## Tarea programada

Configurada con `setup_tarea_asistencia.bat` (en la raíz del proyecto). Se ejecuta de lunes a sábado a las 9:30 AM bajo el usuario `developer7`.

```
Tarea: AsistenciaPlanilla_930
Horario: lun–sáb 09:30
Script: asistencia_planilla/asistencia_planilla.py
```
