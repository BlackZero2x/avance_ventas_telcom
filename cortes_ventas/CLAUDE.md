# Cortes de Ventas — CLAUDE.md

Subproyecto independiente dentro de AVANCE MOVISTAR. Genera y envía informes horarios de ventas declaradas por supervisores vía WhatsApp.

## Qué hace

- Lee respuestas de un Google Sheet (hoja `Respuestas` + hoja `CUOTAS`)
- Calcula tablas por ZONAL y por SUPERVISOR para cada corte horario
- Genera imágenes Excel formateadas (via openpyxl + xlwings)
- Envía las imágenes a grupos de WhatsApp según el corte y el día de la semana

## Cómo ejecutar

```bash
# Informe de corte
python cortes_ventas/cortes_ventas.py --corte 12PM
python cortes_ventas/cortes_ventas.py --corte 2PM
python cortes_ventas/cortes_ventas.py --corte 4PM
python cortes_ventas/cortes_ventas.py --corte 6PM
python cortes_ventas/cortes_ventas.py --corte CIERRE

# Alerta previa (10 min antes del corte)
python cortes_ventas/cortes_ventas.py --corte 12PM --alerta
python cortes_ventas/cortes_ventas.py --corte CIERRE --alerta

# Solo generar imágenes en temp/ sin enviar por WhatsApp (útil para revisar visualmente)
python cortes_ventas/cortes_ventas.py --corte CIERRE --solo-generar
python cortes_ventas/cortes_ventas.py --corte 12PM --solo-generar --fecha 16/06/2026
```

## Rutas importantes

El script vive en `cortes_ventas/cortes_ventas.py` pero usa recursos de la raíz del proyecto:
- `BASE_DIR = Path(__file__).parent.parent` — apunta a la raíz
- `config.json` → `whatsapp_server/config.json`
- `token.json`, `credentials.json` → raíz
- `logs/`, `temp/` → raíz

## Reglas de envío por día

| Día del corte | Cortes 12PM–6PM | CIERRE (ejecuta día siguiente 9AM) |
|---|---|---|
| Lun–Vie | VPA + Supervisores | VPA + Supervisores |
| Sábado | VPA + Supervisores | solo Supervisores |
| Domingo | ningún grupo | ningún grupo |

- **Grupo VPA** (`⚡⚡VPA - Auren`): recibe tabla con filas de región intercaladas (`_insertar_filas_region`)
- **Grupo Supervisores** (`Canal Fija 2026 Supervisores y Jefes`): recibe tabla plana por zonal + tabla por supervisor

## Regiones (tabla VPA)

```python
REGIONES = {
    "REGION LIMA":  ["LIMA"],
    "REGION NORTE": ["CHIMBOTE", "TRUJILLO", "NORTE CHICO"],
    "REGION SUR":   ["AREQUIPA", "ILO", "TACNA"],
}
```

Zonales no listadas aquí aparecen al final de la tabla sin encabezado de región.

## Teléfonos de supervisores

Se leen desde la columna `TELEFONO` de la hoja `CUOTAS` del Google Sheet. Formato aceptado: `9XXXXXXXX` (9 dígitos) o `519XXXXXXXX` (con prefijo 51). Se usan para las menciones en las alertas previas.

## Google Sheets

### Cortes de Ventas
ID en variable de entorno `CORTES_SHEET_ID`. Hojas utilizadas:
- `Respuestas` — registros enviados por supervisores (ZONAL, SUP, CORTE, FECHA_CORTE, VENTA_REGULAR, VENTA_FLEX)
- `CUOTAS` — cuotas diarias por supervisor (ZONAL, SUPERVISOR, CUOTA_DIA, TELEFONO)

### VENTORY — Ventas Registradas
ID en variable de entorno `SHEET_ID_VENTORY` (gid=1543273197). 
La tabla de SUPERVISOR ahora incluye columna `VENTAS_REGISTRADAS_HOY`:
- Contador acumulado de ventas registradas durante el día hasta la hora actual
- Filtros: `DAY = "HOY"`, `Vta_Hoy = "Si"`, `HORA <= hora_actual`
- Ver `VENTORY_INTEGRACIÓN.md` para detalles

## Tareas programadas

Configuradas en `setup_tareas_cortes.bat`. Se ejecutan lunes a sábado:
- Alertas: 11:55, 13:55, 15:55, 17:55, 08:55 (CIERRE)
- Informes: 12:05, 14:05, 16:05, 18:05, 09:05 (CIERRE)
