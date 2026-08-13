# Análisis de Horarios — Envíos WhatsApp en AVANCE MOVISTAR

**Fecha:** 2026-06-20  
**Estado:** 3 subproyectos capturan Excel + envían WhatsApp usando el MISMO servidor  
**Riesgo:** Sin locker de mutex, pueden colisionar en acceso a COM de Excel

---

## Timeline de Horarios

### Lunes a Viernes (Mon–Fri)

```
08:30 AM  AVANCE          Captura TDS (Tabla 1 + Tabla 2) → Grupo JEFES + Jesús
                          [Task: Automatizacion_AVANCE - triggered por email]
                          ├─ jefes_process.py: 2 capturas de rango Excel
                          └─ Envío: 5s + imagen + texto + menciones + 12s + imagen

09:05 AM  CORTES (CIERRE) Informe CIERRE (ejecutado desde día anterior 8:55 AM alerta)
                          [Task: CortesInforme_CIERRE - 09:05]
                          └─ cortes_ventas.py: captura tabla (zonal + supervisor)

09:30 AM  ASISTENCIA      Informe diario PLANILLA (zonal + supervisor)
                          [Task: AsistenciaPlanilla_930 - lun–sáb]
                          └─ asistencia_planilla.py: 2 capturas de rango Excel

11:55 AM  CORTES (ALERTA) Alerta previa (12PM corte)
12:05 PM  CORTES (12PM)   Informe 12PM (zonal + supervisor)
                          [Tasks: CortesAlerta_12PM @11:55 + CortesInforme_12PM @12:05]

13:55 PM  CORTES (ALERTA) Alerta previa (2PM corte)
14:05 PM  CORTES (2PM)    Informe 2PM
                          [Tasks: CortesAlerta_2PM + CortesInforme_2PM]

15:55 PM  CORTES (ALERTA) Alerta previa (4PM corte)
16:05 PM  CORTES (4PM)    Informe 4PM
                          [Tasks: CortesAlerta_4PM + CortesInforme_4PM]

17:55 PM  CORTES (ALERTA) Alerta previa (6PM corte)
18:05 PM  CORTES (6PM)    Informe 6PM
                          [Tasks: CortesAlerta_6PM + CortesInforme_6PM]
```

### Sábado (Saturday)

```
09:05 AM  CORTES (CIERRE) Solo supervisores (sin VPA en cierre sábado)
09:30 AM  ASISTENCIA      Ejecuta normalmente (no hay filtro de día en asistencia_planilla)
11:55 AM  CORTES (ALERTA) Alertas normales (lun–sáb)
12:05 PM  CORTES (12PM)   Informes hasta 18:05
```

### Domingo

```
NO HAY ENVÍOS (cortes deshabilitados, asistencia sin vendedores PLANILLA)
```

---

## Identificación de Colisiones Potenciales

### Escenarios de riesgo alto

| Hora | Proyecto 1 | Proyecto 2 | Riesgo | Lock necesario |
|------|-----------|-----------|--------|----------------|
| **08:30–08:35** | AVANCE (Tabla 1) | — | Bajo | ✓ |
| **08:35–08:40** | AVANCE (Tabla 2) | — | Bajo | ✓ |
| **09:05–09:10** | CORTES CIERRE | ASISTENCIA (09:30 start) | Bajo | ✓ |
| **09:30–09:40** | ASISTENCIA (2 capturas) | — | Bajo | ✓ |
| **11:55–12:10** | CORTES (alerta + informe) | — | Muy Bajo | ✓ |
| **14:05–14:15** | CORTES (2PM) | — | Muy Bajo | ✓ |
| **16:05–16:15** | CORTES (4PM) | — | Muy Bajo | ✓ |
| **18:05–18:15** | CORTES (6PM) | — | Muy Bajo | ✓ |

### Diagnóstico

**Sin locker:**
- Si AVANCE trigger llega fuera de horario (8:40 AM en lugar de 8:30 AM) → **colisión con CORTES CIERRE @ 09:05**
- Si hay reintento de AVANCE/ASISTENCIA por error → **colisión posible**
- Cada proyecto abre COM de Excel, xlwings puede crashear si dos acceden simultáneamente

**Con locker:**
- Cada proyecto aguarda el lock (timeout 30s)
- Orden FIFO garantizado
- Logs de quién está capturando y cuánto tiempo tarda

---

## Implementación del Locker

### Archivos a modificar

1. **Copiar `screenshot_safe.py` a módulos compartidos**
   ```
   C:\proyectos\SSFF\Reportes_ssff_wsp\screenshot_safe.py
   →
   C:\proyectos\AVANCE_MOVISTAR\modules\shared\screenshot_safe.py
   ```

2. **jefes_process.py** — Envolver `_capturar_rango()`
   ```python
   from shared.screenshot_safe import ScreenshotManager
   
   mgr = ScreenshotManager("MOVISTAR_AVANCE_TDS")
   if mgr.adquirir_lock(timeout=30):
       try:
           # ... xlwings + captura ...
       finally:
           mgr.liberar_lock()
   ```

3. **cortes_ventas.py** — Envolver captura
   ```python
   mgr = ScreenshotManager(f"MOVISTAR_CORTE_{corte}")
   if mgr.adquirir_lock(timeout=15):
       try:
           # ... captura Excel ...
       finally:
           mgr.liberar_lock()
   ```

4. **asistencia_planilla.py** — Envolver captura
   ```python
   mgr = ScreenshotManager("MOVISTAR_ASISTENCIA")
   if mgr.adquirir_lock(timeout=30):
       try:
           # ... captura Excel ...
       finally:
           mgr.liberar_lock()
   ```

### IDs únicos recomendados

| Proyecto | ID | Timeout |
|----------|----|---------:|
| AVANCE TDS | `MOVISTAR_AVANCE_TDS` | 30s |
| CORTES 12PM | `MOVISTAR_CORTE_12PM` | 15s |
| CORTES 2PM | `MOVISTAR_CORTE_2PM` | 15s |
| CORTES 4PM | `MOVISTAR_CORTE_4PM` | 15s |
| CORTES 6PM | `MOVISTAR_CORTE_6PM` | 15s |
| CORTES CIERRE | `MOVISTAR_CORTE_CIERRE` | 15s |
| ASISTENCIA | `MOVISTAR_ASISTENCIA` | 30s |

### Directorio de locks (compartido)

```
C:\proyectos\locks\
├─ MOVISTAR_AVANCE_TDS.lock
├─ MOVISTAR_CORTE_12PM.lock
├─ MOVISTAR_ASISTENCIA.lock
└─ (también contiene locks de SSFF)
```

---

## Diagrama de Flujo (Futura)

```
┌─ main_v2.py ─────────┐
│ Detecta trigger 8:30 │
└─────────┬─────────────┘
          ↓
┌─ jefes_process.py ─────────────────────┐
│ 1. Adquirir MOVISTAR_AVANCE_TDS.lock   │
│ 2. Abrir Excel (xlwings)               │
│ 3. Capturar rango B4:AB15              │ (5–10 segundos)
│ 4. Enviar a JEFES + Jesús              │ (5s delay)
│ 5. Capturar rango AC4:BE17             │ (5–10 segundos)
│ 6. Enviar a JEFES + Jesús              │ (12s delay)
│ 7. Liberar MOVISTAR_AVANCE_TDS.lock    │
└─────────┬─────────────────────────────┘
          ↓ (Si cortes intenta capturar 09:05, espera el lock)
┌─ cortes_ventas.py @ 09:05 ─────────────┐
│ 1. Adquirir MOVISTAR_CORTE_CIERRE.lock │ (espera si AVANCE sigue)
│ 2. Capturar tabla (zonal + supervisor) │
│ 3. Enviar a grupos VPA + Supervisores  │
│ 4. Liberar lock                        │
└─────────┬─────────────────────────────┘
          ↓
┌─ asistencia_planilla.py @ 09:30 ──────┐
│ 1. Adquirir MOVISTAR_ASISTENCIA.lock   │
│ 2. Capturar 2 imágenes                 │
│ 3. Enviar a grupo Gestión AUREN        │
│ 4. Liberar lock                        │
└────────────────────────────────────────┘
```

---

## Siguiente: Implementación

Pasos recomendados:
1. ✓ Auditoría completada (archivo actual)
2. ⏳ Copiar `screenshot_safe.py` a `modules/shared/`
3. ⏳ Integrar locker en `jefes_process.py`
4. ⏳ Integrar locker en `cortes_ventas.py`
5. ⏳ Integrar locker en `asistencia_planilla.py`
6. ⏳ Prueba: ejecutar AVANCE + ASISTENCIA simultáneos, verificar logs
7. ⏳ Prueba: ejecutar CORTES solapados, verificar que esperan lock

---

## Referencias

- `screenshot_safe.py`: `C:\proyectos\SSFF\Reportes_ssff_wsp\screenshot_safe.py`
- Shared locks: `C:\proyectos\locks\`
- AVANCE CLAUDE.md: `CLAUDE.md` (en raíz)
- CORTES CLAUDE.md: `cortes_ventas/CLAUDE.md`
- ASISTENCIA CLAUDE.md: `asistencia_planilla/CLAUDE.md`
