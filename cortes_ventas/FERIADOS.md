# Feriados — Deshabilitar cortes automáticos

El script `cortes_ventas.py` se puede configurar para omitir automáticamente los días feriados sin necesidad de modificar el código ni las tareas programadas del Programador de Tareas.

## Configuración

En el archivo `.env` de la raíz del proyecto, agregar la variable:

```env
CORTES_FERIADOS_FECHAS=2026-07-28,2026-07-29,2026-08-30
```

**Formato:**
- Fechas en `YYYY-MM-DD`
- Separadas por coma sin espacios
- Si la lista está vacía o ausente, no hay feriados configurados

## Cómo funciona

1. Al ejecutarse `cortes_ventas.py`, sea manualmente o desde el Programador de Tareas, verifica si la fecha actual (o la especificada con `--fecha`) está en la lista de feriados.
2. Si es feriado: **sale inmediatamente sin error** (exit code 0) y registra un warning en el log.
3. Si NO es feriado: procede normalmente con la generación y envío de cortes.

## Ventajas

- **No requiere modificar el Programador de Tareas** — las tareas pueden seguir configuradas normalmente
- **No genera errores** — el script sale limpiamente, sin reportar fallos
- **Flexible** — cambiar los feriados es tan simple como editar `.env`
- **Documentado** — fácil de entender y mantener

## Ejemplo

Si configuraste `CORTES_FERIADOS_FECHAS=2026-07-28,2026-07-29`:

```bash
# 28 de julio (feriado) → sale sin hacer nada
python cortes_ventas/cortes_ventas.py --corte 12PM
# [WARNING] FERIADO DETECTADO (2026-07-28) — script deshabilitado. No se ejecutará nada.

# 29 de julio (otro feriado) → sale sin hacer nada
python cortes_ventas/cortes_ventas.py --corte CIERRE

# 30 de julio (no es feriado) → procede normalmente
python cortes_ventas/cortes_ventas.py --corte 12PM
# [INFO] Leyendo CUOTAS...
# [INFO] Generando imágenes...
# ...
```
