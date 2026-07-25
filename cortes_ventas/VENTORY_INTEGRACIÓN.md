# Integración de VENTORY — Ventas Registradas

## Qué se agregó

La tabla de SUPERVISOR que se envía al grupo `"Canal Fija 2026 Supervisores y Jefes"` ahora incluye una nueva columna final: **`VENTAS_REGISTRADAS_HOY`**

Esta columna muestra un contador acumulado de ventas registradas en VENTORY durante el día hasta la hora actual.

## Datos de origen

- **Sheet ID**: `SHEET_ID_VENTORY` (variable `.env` o constante en código)
- **Hoja**: `Ventas Registradas` (gid=1543273197)
- **Campos utilizados**:
  - `DAY` — debe ser `"HOY"` (o la fecha actual en DD/MM/YYYY)
  - `SUP` — nombre del supervisor
  - `HORA` — hora en formato "hh" (ej: "10", "14", "18")
  - `Vta_Hoy` — debe ser `"Si"` para contar la venta

## Lógica de cálculo

1. Se lee la hoja de Ventas Registradas de VENTORY
2. Se filtran registros por:
   - `DAY = "HOY"` (o la fecha actual si se especifica)
   - `Vta_Hoy = "Si"`
3. Se extrae la `HORA` del campo correspondiente
4. **Solo se cuentan ventas con HORA ≤ hora actual del sistema**
   - Ejemplo: si son las 14:05, solo se cuentan ventas con HORA ≤ 14
5. Se agrega un contador por supervisor a la tabla de SUPERVISOR

## Coluna en la tabla

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `VENTAS_REGISTRADAS_HOY` | int | Contador acumulado de ventas registradas en VENTORY hasta la hora actual (acumulado durante el día) |

La columna se encuentra **al final de la tabla**, después de `PROYECTADO`.

## Ejemplo de salida

```
ZONAL        | SUP          | 12PM | 2PM | 4PM | 6PM | CIERRE | AVANCE_DIA | CUOTA | %ALCANCE | PROYECTADO | VENTAS_REGISTRADAS_HOY
LIMA         | CARLOS P     | 2    | 3   | 1   | 0   | 0      | 6          | 8     | 75%      | 12         | 5
TRUJILLO     | JUAN PÉREZ   | 1    | 2   | 2   | 1   | 0      | 6          | 6     | 100%     | 8          | 4
```

## Implementación técnica

- Función `_leer_ventory_ventas(service, fecha_hoy)` lee y procesa los datos
- `calcular_tabla_supervisor()` ahora acepta parámetro opcional `service`
- Si no se pasa `service`, la columna aparece con valor 0 para todos

## Notas

- La función es **robusta ante errores**: si VENTORY no está disponible o falta la hoja, registra un warning y continúa con valores 0
- El contador es **acumulativo**: no se reinicia por corte horario, es un total del día hasta ahora
- La columna se incluye en la **fila de totales** como suma de todos los supervisores
