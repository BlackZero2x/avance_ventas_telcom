# Resumen Ejecutivo: Fix de Crashes Puppeteer (2026-06-22)

**Fecha:** 2026-06-22  
**Hora del incidente:** 09:43–15:34 (5h 51m de outage)  
**Causa:** Puppeteer/Chromium en estado corrupto → `Promise was collected` errors  
**Status:** ✅ MITIGADO (validaciones implementadas + watchdog mejorado)

---

## Qué pasó

El servidor WhatsApp colapsó durante la ejecución normal del pipeline AVANCE:
- ✅ AVANCE.py completó exitosamente (09:41–09:43)
- ✅ Todos los módulos (Backs, Jefes, Jesus, Cristian, Guillermo, Carlos, Italo) reportaron OK
- ❌ **Pero ningún mensaje llegó a WhatsApp**
- 🔴 Servidor estuvo caído 5 horas 51 minutos (09:43–15:34)

### Síntoma

Logs de wa_server muestran:
```
[ERROR] Error al enviar texto | {"to":"...","error":"Protocol error (Runtime.callFunctionOn): Promise was collected"}
```

Este error es una **excepción de Puppeteer** (no de WhatsApp). Significa que Chrome se degradó y no respondía a las órdenes del Protocol de Chrome DevTools.

---

## Raíz del problema

### Layer 1: Fallo técnico
Puppeteer/Chromium en wa_server entró en estado corrupto (probablemente por:
- Saturación de memoria (Chrome leyendo muchos archivos)
- Timeout de sesión CDP (Chrome DevTools Protocol)
- Incompatibilidad de versiones)

### Layer 2: Falso positivo en wa_client
wa_client.py detectaba HTTP 200 ✓ pero **no validaba el contenido del JSON**:
```python
if resp.status_code == 200:
    return {"success": True}  # ← Bug: no revisa si JSON contiene {"error": "..."}
```

### Layer 3: Alarma silenciosa
execution_log.json marcó todos los módulos como ✅ OK
- Los módulos intentaron enviar
- wa_server respondió HTTP 200 + JSON error
- wa_client no detectó el error
- main_v2.py registró "todos enviados" falsa

---

## Soluciones Implementadas

### 1. ✅ wa_client.py — Validación robusta de respuestas (HECHO)

**Archivo:** `C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_client.py`  
**Cambio:** Líneas 46–84

**Antes:**
```python
if resp.status_code == 200:
    return {"success": True, "data": resp.json()}
```

**Después:**
```python
if resp.status_code == 200:
    resp_json = resp.json()
    # Validar si el servidor devolvió error en el body
    if resp_json.get("error"):
        # Detectar errores de Puppeteer específicos
        if "Promise was collected" in error_msg or "Protocol error" in error_msg:
            # Esperar más y reintentar (backoff exponencial)
            time.sleep(self.retry_delay * attempt * 2)
            continue
        return {"success": False, "error": error_msg, "status": 200}
    return {"success": True, "data": resp_json}
```

**Impacto:**
- ✅ Ahora detecta errores de Puppeteer automáticamente
- ✅ Reintenta con backoff exponencial
- ✅ Si sigue fallando, reporta error real en logs

### 2. ✅ wa_server.js — Watchdog mejorado (HECHO)

**Archivo:** `C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_server.js`  
**Cambios:** Líneas 150–198

**Mejoras:**
- Nueva función `_isPuppeteerCrash()` para detectar:
  - `"Promise was collected"`
  - `"Protocol error"`
  - `"Runtime.callFunctionOn"` errors
- Contador de crashes con reinicio automático en 3 ocurrencias
- Reset de contador si no hay errores durante 10 minutos
- Logs mejorados para debugging

**Impacto:**
- ✅ Ahora el watchdog **detecta Puppeteer crashes proactivamente** (no solo detached Frame)
- ✅ Auto-reinicia Chrome en lugar de esperar 5 minutos
- ✅ Evita cascadas: un crash dispara reinicio antes del siguiente envío

### 3. 📊 Documentación del diagnóstico

**Archivo:** `C:\proyectos\AVANCE_MOVISTAR\DIAGNOSTICO_PUPPETEER_CRASHES.md`

- Análisis completo del incidente
- Cronología detallada
- Causas raíz (3 hipótesis)
- Recomendaciones de largo plazo

---

## Recomendaciones Adicionales (To-Do)

### Urgentes (esta semana)
1. **Actualizar whatsapp-web.js a versión más reciente**
   - Actual: `^1.26.0` (de ~2024-01)
   - Verificar: ¿Existe 1.27.0, 1.28.0 con fixes de Puppeteer?
   - Comando: `cd whatsapp_server && npm update whatsapp-web.js`

2. **Limpieza de memoria en wa_server**
   - Después de cada 50–100 envíos, destruir y recrear Chrome session
   - O agregar `--max-old-space-size=2048` a Node.js

3. **Prueba de carga del servidor**
   - Simular 100+ envíos seguidos para verificar degradación de memoria
   - Verificar que wa_client reintentar automático funcione

### Corto plazo (próximas 2 semanas)
4. **Implementar locker de mutex**
   - Ya documentado en ANALISIS_HORARIOS_WHATSAPP.md
   - Evita solapamiento de capturas Excel (AVANCE, CORTES, ASISTENCIA)

5. **Monitoring y alertas**
   - Dashboard de salud del wa_server (uptime, crashes, memory)
   - Alertas en Slack si wa_server devuelve `status: not_ready`

### Monitoreo continuo
6. **Logs centralizados**
   - Agregar timestamp más granular en wa_server logs
   - Correlacionar errores de wa_server con envíos de wa_client

---

## Verificación

### Test inmediato (ahora)
```bash
# Verificar que wa_client.py está actualizado
grep -n "Promise was collected" C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_client.py

# Verificar que wa_server.js está actualizado
grep -n "_isPuppeteerCrash" C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_server.js
```

### Test de recuperación (próximo trigger AVANCE)
- Monitorear logs durante ejecución
- Si wa_server crashea, verificar que wa_client reintentar automático funcione
- Si no llega mensaje, revisar: `whatsapp_server/logs/wa_server_2026-06-23.log`

---

## Impact

### Antes del fix
- ❌ Crash de Puppeteer → silencio de 5 horas
- ❌ execution_log.json marca "todo OK" falso
- ❌ Usuarios nunca se enteran de que falló

### Después del fix
- ✅ wa_client detecta error automáticamente
- ✅ Reintenta con backoff exponencial
- ✅ wa_server watchdog reinicia Chrome proactivamente
- ✅ Logs claros de qué falló y por qué

---

## Archivos Actualizados

| Archivo | Cambios | Status |
|---------|---------|--------|
| `wa_client.py` | Validación robusta de respuestas + retry backoff | ✅ LISTO |
| `wa_server.js` | Watchdog mejorado para detectar Puppeteer crashes | ✅ LISTO |
| `DIAGNOSTICO_PUPPETEER_CRASHES.md` | Análisis completo del incidente | ✅ LISTO |

---

## Próximos Pasos

1. ⏳ **Verificar que wa_server restartió correctamente** (ya hecho a las 15:34)
2. ⏳ **Esperar siguiente trigger AVANCE** para ver si retry backoff funciona
3. ⏳ **Revisar versiones de whatsapp-web.js** por actualizaciones
4. ⏳ **Implementar locker de mutex** (independiente, pero complementario)

---

## Contactos y Alertas

Si vuelve a ocurrir:
- ✅ wa_client reintentará automáticamente (3 intentos máximo)
- ⚠️ Si sigue fallando → revisar wa_server logs para `Protocol error` o `Promise was collected`
- 🔴 Si wa_server no reinicia → reiniciar manualmente: `taskkill /f /im node.exe`

**Monitoreo recomendado:**
- Revisar `execution_log.json` cada mañana tras AVANCE
- Si ves módulos con `"false"`, revisar `wa_server_YYYY-MM-DD.log`
- Alertar si wa_server estuvo down >1 hora

---

**Resumen final:** La cascada de fallos ha sido mitigada con validaciones robustas en wa_client y watchdog mejorado en wa_server. Próximas mejoras: actualizar dependencias, implementar locker, y agregar alertas.
