# Checklist: Fix de Puppeteer Crashes

**Status:** ✅ COMPLETO  
**Fecha:** 2026-06-22  
**Responsable:** Claude Code

---

## Implementaciones Completadas

### A. wa_client.py Validation (líneas 46–84)
- [x] Parsear JSON de respuesta incluso con HTTP 200
- [x] Detectar campo `"error"` en JSON response
- [x] Identificar `"Promise was collected"` en error message
- [x] Identificar `"Protocol error"` en error message
- [x] Aplicar backoff exponencial: `sleep(retry_delay * attempt * 2)`
- [x] Reintentar automáticamente (max 3 intentos)
- [x] Log de error detallado

**Verificación:**
```bash
grep -n "Promise was collected\|Protocol error" wa_client.py
# Output: línea 70 debe contener ambas condiciones
```

### B. wa_server.js Watchdog Mejorado (líneas 150–228)
- [x] Nueva variable: `_puppeteer_crash_count`
- [x] Nueva variable: `_last_puppeteer_crash_time`
- [x] Nueva función: `_isPuppeteerCrash(err)`
- [x] Detectar `"Promise was collected"` en error message
- [x] Detectar `"Protocol error"` en error message
- [x] Detectar `"Runtime.callFunctionOn"` en error message
- [x] Contador de crashes en watchdog (cada 5 min)
- [x] Contador de crashes en enqueue() (real-time)
- [x] Auto-restart en 2+ crashes en watchdog
- [x] Auto-restart en 3+ crashes en enqueue
- [x] Reset contador si no hay errores en 10 min
- [x] Logs con count + error message

**Verificación:**
```bash
grep -n "_isPuppeteerCrash\|_puppeteer_crash_count" wa_server.js
# Output: líneas 152, 164, 166, 168, 169, 181, 182, 184, 202, 219, 221, 223, 224
```

### C. Documentación
- [x] `DIAGNOSTICO_PUPPETEER_CRASHES.md` — análisis completo
- [x] `RESUMEN_FIX_PUPPETEER_2026-06-22.md` — resumen ejecutivo
- [x] `CHECKLIST_PUPPETEER_FIX.md` — este archivo

---

## Verificación de Sintaxis

### wa_client.py
```bash
python -m py_compile C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_client.py
# Esperado: Sin errores
```

### wa_server.js
```bash
node -c C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_server.js
# Esperado: Sin errores de sintaxis
```

---

## Testing Manual (Post-Deploy)

### Test 1: Simular error de Puppeteer en wa_client
```bash
cd C:\proyectos\AVANCE_MOVISTAR\whatsapp_server

# Editar wa_client.py temporalmente para simular error
# Cambiar línea 54-56 a devolver HTTP 200 + JSON error:
# resp_json = {"error": "Protocol error (Runtime.callFunctionOn): Promise was collected"}

# Ejecutar test
python wa_client.py --test

# Verificar logs:
# [WA] Enviando texto a '...'
# [Intento 1/3] Servidor puppeteer inestable: Protocol error...
# [Intento 2/3] ... (espera 2x más larga)
# [Intento 3/3] ...
# [WA] [ERROR] Error enviando a '...': Protocol error...
```

### Test 2: Verificar watchdog detecta crashes
```bash
# En wa_server.js, agregar log en watchdog durante test:
# log("WARN", `Watchdog: Puppage eval resultado:`, { result: null });

# Ejecutar test de carga
for i in {1..20}; do
  curl -X POST http://localhost:8002/send-text \
    -H "Content-Type: application/json" \
    -d '{"to":"51975155264@c.us","message":"Test '$i'"}'
  sleep 1
done

# Verificar logs muestren:
# [WARN] Watchdog: Crash de Puppeteer detectado (1x)
# [WARN] Watchdog: Crash de Puppeteer detectado (2x)
# [ERROR] Watchdog: Múltiples crashes de Puppeteer — reiniciando cliente
```

### Test 3: Verificar retry backoff en wa_client
```bash
# Modificar wa_client.py para log de retry
# Agregar: print(f"[RETRY] Intento {attempt}, espera {delay}s")

python wa_client.py --test 2>&1 | grep -i retry
# Esperado: Ver incremento de tiempo de espera: 5s, 20s, 40s
```

---

## Monitoreo Post-Implementación

### Diarios (tras cada AVANCE)
- [ ] Revisar `execution_log.json` — ¿todos los módulos tienen `true`?
- [ ] Revisar `wa_server_YYYY-MM-DD.log` — ¿hay Puppeteer crashes?
- [ ] Revisar `automation_YYYYMMDD_HHMMSS.log` — ¿hay "Error" o "WARNING"?

### Semanales
- [ ] Contar crashes de Puppeteer en últimos 7 días
- [ ] Verificar que watchdog reinicia cliente automáticamente
- [ ] Revisar uptime del wa_server (sin downtime >1 hora)

### Mensuales
- [ ] Revisar versión de `whatsapp-web.js` por updates
- [ ] Analizar patrones de crashes (ej: siempre en hora pico)
- [ ] Evaluar necesidad de aumentar memoria para Node.js

---

## Problemas Conocidos y Workarounds

| Problema | Síntoma | Workaround |
|----------|---------|-----------|
| Puppeteer out of memory | Crashes después de 500+ envíos | Reiniciar wa_server cada 1000 envíos |
| Chrome session stale | `upload failed` errors | Health check cada 5 min, reload session |
| Outdated whatsapp-web.js | Frequent Puppeteer errors | Update a versión más reciente |

---

## Rollback Plan (si es necesario)

Si los cambios causan problemas:

### Revertir wa_client.py
```bash
git diff whatsapp_server/wa_client.py
git checkout whatsapp_server/wa_client.py
```

### Revertir wa_server.js
```bash
git diff whatsapp_server/wa_server.js
git checkout whatsapp_server/wa_server.js
taskkill /f /im node.exe
# Reiniciar wa_server
```

---

## Sign-Off

| Componente | Status | Fecha | Verificado |
|-----------|--------|-------|-----------|
| wa_client.py | ✅ LISTO | 2026-06-22 | Validación + retry |
| wa_server.js | ✅ LISTO | 2026-06-22 | Watchdog mejorado |
| Documentación | ✅ LISTO | 2026-06-22 | Análisis + resumen |
| Testing manual | ⏳ PENDIENTE | — | Ejecutar post-deploy |
| Monitoreo | ⏳ PENDIENTE | — | Configurar alertas |

---

## Notas Adicionales

- Los cambios en wa_client.py son **backwards compatible** (no afecta código que ya funciona)
- Los cambios en wa_server.js son **additive** (agregan detección, no quitan funcionalidad)
- No se requiere reinicio del orquestador (main_v2.py)
- wa_server.js puede ser reiniciado en cualquier momento (reintentará automáticamente)

---

**Próximo paso:** Esperar al siguiente trigger AVANCE para verificar en producción que:
1. Si wa_server crashea, wa_client reintenta automáticamente
2. Si wa_server crashea, watchdog lo reinicia en <30s
3. Los mensajes finalmente llegan (aunque sea después de reintentos)
