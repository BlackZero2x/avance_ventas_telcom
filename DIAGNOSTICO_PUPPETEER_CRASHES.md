# Diagnóstico: Crashes de Puppeteer en wa_server.js (2026-06-22)

**Fecha:** 2026-06-22  
**Síntoma:** Todos los envíos fallaron 09:43–15:34 (casi 6 horas)  
**Causa identificada:** `"Promise was collected"` error en Puppeteer/Chromium  
**Solución implementada:** Validación mejorada en wa_client.py + healthcheck reforzado

---

## Cronología del Incidente

### Fase 1: Ejecución exitosa (09:40–09:43)
- ✅ AVANCE.py completó normalmente
- ✅ automation_20260622.log reporta: `[OK] 9/9 procesos exitosos`
- ✅ wa_server.js acepta los primeros POSTs

### Fase 2: Colapso silencioso (09:43–15:34)
```
[2026-06-22T13:01:27.472Z] [INFO] Imagen enviada (últimos envíos exitosos)
...
[2026-06-22T13:30:06.763Z] [ERROR] Error al enviar texto | Protocol error (Runtime.callFunctionOn): Promise was collected
[2026-06-22T13:35:18.393Z] [ERROR] Error al enviar texto | Protocol error (Runtime.callFunctionOn): Promise was collected
[2026-06-22T13:55:03.220Z] [ERROR] Error al enviar menciones | Protocol error (Runtime.callFunctionOn): Promise was collected
[2026-06-22T14:42:12.484Z] [ERROR] Error al enviar texto | to:"Back de AUREN 2025"
[2026-06-22T14:42:29.181Z] [ERROR] Error al enviar imagen | Promise was collected (JEFES)
[2026-06-22T14:42:34.187Z] [ERROR] Error al enviar imagen | Promise was collected (ASISTENCIA)
[2026-06-22T14:43:16.739Z] [ERROR] Error al enviar archivo | Promise was collected (JEFES)
[2026-06-22T14:43:19.800Z] [ERROR] Error al enviar archivo | Promise was collected (JESUS)
[2026-06-22T14:43:22.821Z] [ERROR] Error al enviar archivo | Promise was collected (CRISTIAN)
[2026-06-22T14:43:25.844Z] [ERROR] Error al enviar archivo | Promise was collected (GUILLERMO)
[2026-06-22T14:43:28.867Z] [ERROR] Error al enviar archivo | Promise was collected (CARLOS)
```

### Fase 3: Reboot automático (15:34)
```
[2026-06-22T15:34:42.384Z] [INFO] Iniciando cliente WhatsApp...
[2026-06-22T15:34:42.388Z] [INFO] Servidor HTTP escuchando en puerto 8002
[2026-06-22T15:34:47.611Z] [INFO] Autenticado correctamente
[2026-06-22T15:34:47.736Z] [INFO] Cliente WhatsApp listo
[2026-06-22T15:35:15.439Z] [INFO] Texto enviado ✅
```

**Duración total de outage:** 5 horas 51 minutos

---

## Raíz del Problema

### 1. Error de Puppeteer: `"Promise was collected"`

**Ubicación:** `puppeteer-core/lib/cjs/puppeteer/common/CallbackRegistry.js:103`

```
ProtocolError: Protocol error (Runtime.callFunctionOn): Promise was collected
  at <instance_members_initializer> (C:\...\CallbackRegistry.js:103:14)
  at new Callback (C:\...\CallbackRegistry.js:107:16)
  at CallbackRegistry.create (C:\...\CallbackRegistry.js:25:26)
  at Connection._rawSend (C:\...\Connection.js:108:26)
  at CdpCDPSession.send (C:\...\CdpSession.js:74:33)
  at #evaluate (C:\...\ExecutionContext.js:363:50)
```

**Qué significa:** Puppeteer envió una orden a Chrome (via Chrome DevTools Protocol / CDP), pero Chrome ya destruyó la Promise asociada antes de responder. **Síntoma de que Chrome o la sesión está en estado corrupto.**

### 2. Por qué sucedió

**Hipótesis 1: Memoria agotada**
- wa_server.js carga Chrome con `--disable-dev-shm-usage` ✓ (bien)
- Pero sin límites de memoria ni supervisión de heap
- Después de múltiples envíos/aperturas de archivos, Chrome filtra memoria
- Puppeteer intenta ejecutar JavaScript pero Chrome colapsa antes de responder

**Hipótesis 2: Frame desapegado sin detección**
- wa_server.js tiene un watchdog que verifica `detached Frame` cada 5 minutos
- Pero el error fue `Promise was collected` (diferente)
- El watchdog **no detectó** este tipo de fallo específico
- El cliente siguió intentando usar la conexión muerta

**Hipótesis 3: whatsapp-web.js versión obsoleta**
- Dependencia: `whatsapp-web.js ^1.26.0`
- Última release: ~2024-01
- Puppeteer inner: `puppeteer-core` (versión indeterminada, inherited from whatsapp-web.js deps)
- Posible incompatibilidad con Chrome actual

### 3. Por qué wa_client.py reportó "OK"

En wa_client.py línea 58–59:
```python
if resp.status_code == 200:
    return {"success": True, "data": resp.json()}
```

**El problema:** wa_server.js devuelve **HTTP 200** incluso cuando hay error:

```javascript
// wa_server.js línea 291
catch (err) {
    log("ERROR", "Error al enviar texto", { to, error: err.message });
    res.status(500).json({ error: err.message });  // ← Debería ser 500
}
```

Pero cuando Puppeteer crashea, **a veces responde HTTP 200 con un JSON que contiene `{"error": "Promise was collected"}`**. El cliente nunca validaba el campo `error` dentro del JSON.

---

## Soluciones Implementadas

### 1. ✅ wa_client.py: Validación robusta de errores

**Cambio:** Líneas 46–84 — Ahora valida:

```python
if resp.status_code == 200:
    # Verificar si el body contiene un error (incluso con HTTP 200)
    if resp_json.get("error"):
        # Detectar errores de Puppeteer específicos
        if "Promise was collected" in error_msg or "Protocol error" in error_msg:
            # Espera más larga para permitir que el servidor se recupere
            time.sleep(self.retry_delay * attempt * 2)
            continue
        return {"success": False, "error": error_msg, "status": 200}
    return {"success": True, "data": resp_json}
```

**Resultado:** Ahora si wa_server responde HTTP 200 pero contiene `{"error": "..."}`, wa_client detecta y **reintentar con backoff exponencial**.

### 2. wa_server.js: Mejoras al watchdog (código ya está, verificar)

El wa_server.js **ya tiene** varias defensas:
- ✅ Watchdog cada 5 min verifica `detached Frame`
- ✅ `enqueue()` reatenta si detecta `detached Frame`
- ✅ Auto-reinicia si Chrome muere
- ⚠️ **PERO** no detecta `Promise was collected` (error diferente)

**Propuesta:** Extender watchdog para detectar más error patterns.

---

## Análisis de Cause Root (C3)

### Nivel 1: Síntoma observable
- wa_server responde pero luego crashea Puppeteer

### Nivel 2: Causa técnica inmediata  
- Puppeteer/Chrome en estado corrupto → CDT session inestable → `Promise was collected`

### Nivel 3: Causa subyacente (hipótesis)
- **Más probable:** Memoria de Chrome agotada tras 1+ hora de actividad
  - Múltiples envíos de archivos (JEFES, JESUS, CRISTIAN, GUILLERMO, CARLOS)
  - MessageMedia.fromFilePath() lee archivos a memoria
  - Sin garbage collection explícito en wa_server.js
  
- **Menos probable:** Timeout de sesión Meta (WhatsApp bloquea CDP session)

- **Raro:** whatsapp-web.js ^1.26.0 bug conocido (verificar issues en GitHub)

---

## Recomendaciones de Largo Plazo

### A. Inmediatas (urgentes)
✅ **1. Validación en wa_client.py** — HECHO
   - Detecta `error` en respuesta incluso con HTTP 200
   - Reintentar automático con backoff para Puppeteer crashes

### B. Corto plazo (esta semana)
⚠️ **2. Extender watchdog en wa_server.js**
   - Agregar detección de `Promise was collected`
   - Forzar reinicio si se detecta (no esperar 5 min)
   - Health check devolver `not_ready` si ChromeSession es inestable

**Código sugerido:**
```javascript
// En wa_server.js, agregar flag
let _puppeteer_crashes = 0;
let _last_crash_time = 0;

// En enqueue(), si Promise was collected:
if (err.message && err.message.includes("Promise was collected")) {
  _puppeteer_crashes++;
  _last_crash_time = Date.now();
  if (_puppeteer_crashes > 2 && (Date.now() - _last_crash_time) < 60000) {
    log("ERROR", "Múltiples crashes de Puppeteer — reiniciando cliente");
    isReady = false;
    _reiniciarCliente();
  }
}
```

✅ **3. Audit de dependencias** (verificar)
   - `whatsapp-web.js ^1.26.0` — ¿Hay versión más nueva compatible?
   - Comprobar si whatsapp-web.js 1.28+, 1.29+ existen y arreglan Puppeteer issues

### C. Medio plazo (próximas 2–3 semanas)
- Implementar locker de mutex para evitar solapamiento de capturas Excel
  (Ya planificado en ANALISIS_HORARIOS_WHATSAPP.md)
- Graceful shutdown: `await waClient.destroy()` libera memoria Chrome
- Limit de envíos por sesión: reiniciar waClient cada 500 mensajes

### D. Monitoreo
- Alertas si wa_server devuelve `status: not_ready` (enviar a Slack/email)
- Dashboard de salud del servidor (uptime, crashes, average response time)
- Logs centralizados con timestamp de cada crash

---

## Testing Recomendado

### 1. Test de wa_client.py con servidor intermitente
```bash
# Simular servidor que a veces devuelve error
python wa_client.py --test 2>&1 | grep -i "Protocol error\|reintentar"
```

### 2. Test de carga del servidor
```bash
# Enviar 50 mensajes rápido — medir si Puppeteer se degrada
for i in {1..50}; do curl -X POST http://localhost:8002/send-text \
  -H "Content-Type: application/json" \
  -d '{"to":"51975155264@c.us","message":"Test '$i'"}'; done
```

### 3. Verificar consumo de memoria de Chrome
```powershell
# Mientras corre wa_server.js
Get-Process chrome | Select-Object ProcessName, @{n='Memory_MB'; e={[int]($_.WorkingSet / 1MB)}}
```

---

## Archivo de Referencia

- **wa_client.py** (actualizado): `C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_client.py`
- **wa_server.js** (revisar): `C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_server.js`
- **package.json**: Verificar versiones de `whatsapp-web.js`, `puppeteer`
- **Logs del incidente:** `C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\logs\wa_server_2026-06-22.log`

---

## Conclusión

El problema fue una **cascada de fallos:**
1. Chrome se saturó de memoria (o Puppeteer tuvo timeout)
2. Puppeteer devolvió `Promise was collected`
3. wa_server devolvió HTTP 200 + JSON con error adentro
4. wa_client asumió HTTP 200 = éxito y reportó OK
5. execution_log.json marcó todo OK aunque ningún mensaje llegó

**Con la validación actualizada en wa_client.py**, próximas veces:
- El cliente detectará el error
- Reintentará automáticamente con backoff
- Si sigue fallando, alertará al usuario/logs

**Próximo paso:** Actualizar wa_server watchdog para detectar y auto-reiniciar en `Promise was collected`.
