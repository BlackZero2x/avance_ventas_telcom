const { Client, LocalAuth, MessageMedia } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");
const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");
require("dotenv").config({ path: path.join(__dirname, "..", ".env") });

// ============================================================
// CONFIGURACIÓN
// ============================================================
const PORT = 8002;
const SESSION_DIR = path.join(__dirname, "session_data");
const LOG_DIR = path.join(__dirname, "logs");
const CONFIG_PATH = path.join(__dirname, "config.json");

[SESSION_DIR, LOG_DIR].forEach((dir) => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// ============================================================
// LOGGER
// ============================================================
function log(level, message, data) {
  const timestamp = new Date().toISOString();
  const entry = `[${timestamp}] [${level}] ${message}${data ? " | " + JSON.stringify(data) : ""}`;
  console.log(entry);
  const today = new Date().toISOString().slice(0, 10);
  const logFile = path.join(LOG_DIR, `wa_server_${today}.log`);
  fs.appendFileSync(logFile, entry + "\n");
}

// ============================================================
// CARGAR CONFIG
// ============================================================
function loadConfig() {
  try {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf-8"));
  } catch {
    log("WARN", "No se pudo cargar config.json, usando config vacia");
    return { groups: {}, contacts: {} };
  }
}

// ============================================================
// RESOLVER DESTINATARIO
// ============================================================
function resolveRecipient(nameOrId) {
  if (nameOrId.includes("@")) return nameOrId;
  const config = loadConfig();
  if (config.groups && config.groups[nameOrId]) return config.groups[nameOrId];
  if (config.contacts && config.contacts[nameOrId]) return config.contacts[nameOrId];
  return null;
}

// ============================================================
// CLIENTE WHATSAPP
// ============================================================
let waClient = null;
let isReady = false;
let _lastUploadError = null;   // timestamp del último "upload failed" para el health check
let _lastNotificacion = null;  // timestamp del último correo de alerta (throttle 30 min)

// Envía una alerta por correo al administrador cuando la sesión WA falla.
// Tiene throttle: solo envía un correo cada 30 minutos aunque el error persista.
function _notificarError(tipo, detalle) {
  const THROTTLE_MS = 30 * 60 * 1000; // 30 minutos
  const ahora = Date.now();
  if (_lastNotificacion && (ahora - _lastNotificacion) < THROTTLE_MS) {
    log("INFO", "Notificacion de error omitida (throttle 30 min activo)");
    return;
  }
  _lastNotificacion = ahora;

  const { exec } = require("child_process");
  const pythonExe = "C:\\proyectos\\.venv\\Scripts\\python.exe";
  const script = path.join(__dirname, "notify_error.py");
  const safeDetalle = (detalle || "").replace(/"/g, "'").replace(/\n/g, " ");
  const cmd = `"${pythonExe}" "${script}" --tipo "${tipo}" --detalle "${safeDetalle}"`;

  log("INFO", "Enviando alerta por correo", { tipo });
  exec(cmd, { cwd: path.join(__dirname, "..") }, (err, stdout, stderr) => {
    if (err) {
      log("WARN", "No se pudo enviar alerta por correo", { error: err.message, stderr });
    } else {
      log("INFO", "Alerta por correo enviada", { stdout: stdout.trim() });
    }
  });
}

function startWhatsApp() {
  log("INFO", "Iniciando cliente WhatsApp...");

  // Resolver la version de WA Web mas reciente disponible en cache local.
  // La libreria v1.34.7 pide por defecto la 2.3000.1017054665, que no esta en
  // nuestro cache (el mas reciente es del 08/08/2026). Si esa version no existe
  // localmente Y el proxy bloquea web.whatsapp.com en la descarga de la version
  // exacta, el cliente se queda colgado entre "authenticated" y "ready"
  // indefinidamente (authTimeoutMs=0 => sin timeout). Fix: inyectar la version
  // mas reciente del cache local para que no haga request externo. (08/08/2026)
  const _wwebCacheDir = path.join(__dirname, ".wwebjs_cache");
  let _wwebVersion = "2.3000.1017054665"; // fallback al default de la libreria
  try {
    const _cacheFiles = fs.readdirSync(_wwebCacheDir)
      .filter(f => f.endsWith(".html"))
      .sort(); // orden lexicografico = orden cronologico para este esquema de nombres
    if (_cacheFiles.length > 0) {
      _wwebVersion = _cacheFiles[_cacheFiles.length - 1].replace(".html", "");
      log("INFO", "Usando version WA Web del cache local", { version: _wwebVersion });
    }
  } catch (_) {
    log("WARN", "No se pudo leer cache local de WA Web, usando version default", { version: _wwebVersion });
  }

  waClient = new Client({
    authStrategy: new LocalAuth({
      clientId: "automation_session",
      dataPath: SESSION_DIR,
    }),
    webVersion: _wwebVersion,
    webVersionCache: {
      type: "local",
      path: _wwebCacheDir,
    },
    puppeteer: {
      headless: true,
      executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-extensions",
        "--disable-default-apps",
        "--no-first-run",
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding",
        "--disable-features=TranslateUI",
        // NOTA: se quitaron --single-process y --no-zygote (18/07/2026).
        // Bajo carga sostenida (SSFF cada hora + AVANCE + PBI + VPN) provocaban
        // que Chrome descartara el contexto JS de la pagina => "Promise was collected".
        // Multi-proceso aisla el renderer y evita esa corrupcion de contexto.
        ...(process.env.HTTP_PROXY ? [`--proxy-server=${process.env.HTTP_PROXY}`] : []),
      ],
    },
  });

  // Autenticar proxy antes de cada navegación
  if (process.env.HTTP_PROXY && process.env.PROXY_USER && process.env.PROXY_PASS) {
    const _origInit = waClient.initialize.bind(waClient);
    waClient.initialize = async function () {
      const result = _origInit();
      // Esperar a que pupPage esté disponible y configurar autenticación
      const interval = setInterval(async () => {
        if (waClient.pupPage) {
          clearInterval(interval);
          try {
            await waClient.pupPage.authenticate({
              username: process.env.PROXY_USER,
              password: process.env.PROXY_PASS,
            });
            log("INFO", "Proxy autenticado en pupPage");
          } catch (e) {
            log("WARN", "Error autenticando proxy", { err: e.message });
          }
        }
      }, 200);
      return result;
    };
  }

  waClient.on("qr", (qr) => {
    log("INFO", "QR recibido — escanea con tu WhatsApp:");
    qrcode.generate(qr, { small: true });
  });

  waClient.on("ready", () => {
    isReady = true;
    log("INFO", "Cliente WhatsApp listo");
  });

  waClient.on("authenticated", () => {
    log("INFO", "Autenticado correctamente");
  });

  waClient.on("auth_failure", (msg) => {
    isReady = false;
    log("ERROR", "Fallo de autenticacion — reintentando en 30s", { msg });
    setTimeout(() => _reiniciarCliente(), 30000);
  });

  waClient.on("disconnected", (reason) => {
    isReady = false;
    log("WARN", "Cliente desconectado — reintentando en 20s", { reason });
    _notificarError("Cliente desconectado", `WhatsApp Web se desconecto del servidor. Razon: ${reason}. Reconectando en 20 segundos.`);
    setTimeout(() => _reiniciarCliente(), 20000);
  });

  waClient.initialize();
}

let _reconectando = false;

async function _reiniciarCliente() {
  if (_reconectando) {
    log("INFO", "Reconexion ya en curso, ignorando duplicado");
    return;
  }
  _reconectando = true;
  isReady = false;
  log("INFO", "Destruyendo cliente anterior...");
  try {
    // Timeout de 15s para destroy — si Chrome está completamente bloqueado, no esperar indefinidamente
    if (waClient) {
      await Promise.race([
        waClient.destroy(),
        new Promise((_, rej) => setTimeout(() => rej(new Error("destroy timeout")), 15000)),
      ]);
    }
  } catch (e) {
    log("WARN", "Error al destruir cliente (ignorado)", { e: e.message });
  }
  // Siempre matar Chrome al reiniciar — libera RAM acumulada que causa "Promise was collected"
  try {
    execSync("taskkill /F /IM chrome.exe /T", { stdio: "ignore" });
    log("INFO", "Procesos Chrome terminados forzosamente");
  } catch (_) { /* ignorar si no hay Chrome corriendo */ }

  waClient = null;

  // Esperar 8s para que el SO libere la RAM de Chrome antes de relanzar
  log("INFO", "Esperando 8s para liberar RAM antes de relanzar...");
  await new Promise((res) => setTimeout(res, 8000));

  log("INFO", "Relanzando cliente WhatsApp...");
  _reconectando = false;
  startWhatsApp();
}

// Watchdog: cada 2 minutos verifica que la página de WhatsApp Web sigue operativa.
// Usa un evaluate real sobre el contexto de WA (no solo `true`) para detectar
// Promise was collected antes de que llegue un envío real.
setInterval(async () => {
  if (!isReady || _reconectando) return;
  try {
    // Evalúa algo real en el contexto de la página WA — si el contexto JS fue descartado
    // por Chrome, esto lanzará "Promise was collected" igual que sendMessage
    await waClient.pupPage.evaluate(() => typeof window !== "undefined");
  } catch (e) {
    if (_isDetachedFrame(e)) {
      log("ERROR", "Watchdog: Chrome caido (detached Frame) — reiniciando cliente");
      isReady = false;
      _reiniciarCliente();
    } else if (_isPromiseCollected(e)) {
      log("ERROR", "Watchdog: contexto JS descartado (Promise collected) — reiniciando cliente");
      isReady = false;
      _lastUploadError = Date.now();
      _notificarError("Promise was collected (Watchdog)", "El watchdog de 2 minutos detecto que Chrome perdio el contexto JS de la pagina WA. Reiniciando sesion automaticamente.");
      _reiniciarCliente();
    } else {
      log("WARN", "Watchdog: error inesperado en ping", { error: e.message });
    }
  }
}, 2 * 60 * 1000);

// ============================================================
// COLA DE ENVÍO GLOBAL — compartida por TODOS los proyectos
// (AVANCE_MOVISTAR, SSFF, PBI, VPN_MIFIBRA, tablas Jesús).
//
// Diseño (reescrito 18/07/2026):
//  - Cola explícita (array), no cadena de promesas, para poder
//    DRENARLA de golpe cuando Chrome pierde el contexto y para
//    soportar PRIORIDAD.
//  - Mínimo 5 s entre mensajes salientes (anti-ban).
//  - Prioridad: los envíos con priority alto se atienden primero.
//    AVANCE (evento puntual por trigger de correo) usa priority 10
//    para pasar delante de los cortes horarios de SSFF (priority 0).
//  - Al detectar contexto muerto (Promise collected / Protocol
//    error / detached Frame) se DRENA la cola: todos los pendientes
//    se rechazan al instante en vez de esperar 5 s por cada uno
//    (esto causaba las ~3 h de "cola llena" del 17/07).
// ============================================================
const SEND_DELAY_MS = 5000;
const QUEUE_MAX = 12;          // máximo de mensajes esperando en cola (5 proyectos)
const PRIORITY_AVANCE = 10;    // eventos puntuales (trigger de correo)
const PRIORITY_NORMAL = 0;     // cortes horarios, capturas periódicas

let _queue = [];               // items: { fn, priority, seq, resolve, reject }
let _queueDepth = 0;           // = _queue.length (contador cacheado para los guards)
let _draining = false;         // true mientras el worker procesa un envío
let _seqCounter = 0;           // desempate FIFO dentro de la misma prioridad

function _isDetachedFrame(err) {
  return err && err.message && err.message.includes("detached Frame");
}

function _isUploadFailure(err) {
  return err && err.message && err.message.includes("upload failed");
}

function _isPromiseCollected(err) {
  return err && err.message && err.message.includes("Promise was collected");
}

function _isProtocolError(err) {
  return err && err.message && err.message.includes("Protocol error");
}

function _isUndefinedResult(err) {
  return err && err.message && err.message.includes("Cannot read properties of undefined");
}

// Un contexto muerto justifica drenar la cola completa (todos los
// envíos pendientes fallarían igual, uno por uno, esperando 5 s cada vez).
function _isContextoMuerto(err) {
  return _isDetachedFrame(err) || _isPromiseCollected(err) ||
         _isProtocolError(err) || _isUndefinedResult(err);
}

// Devuelve true si la cola está llena. Los endpoints deben rechazar con 503 si es así.
function queueFull() {
  return _queueDepth >= QUEUE_MAX;
}

// Rechaza al instante todos los envíos pendientes. Se llama cuando el
// contexto de Chrome murió: no tiene sentido esperar 5 s por cada uno.
function _drenarCola(motivo) {
  if (_queue.length === 0) return;
  const n = _queue.length;
  log("WARN", `Drenando cola — ${n} envios pendientes rechazados al instante`, { motivo });
  const pendientes = _queue;
  _queue = [];
  _queueDepth = 0;
  for (const item of pendientes) {
    item.reject(new Error(`Envio cancelado: ${motivo}. Reintenta cuando el servidor reconecte.`));
  }
}

// Encola un envío. priority alto => se atiende antes (dentro de la misma
// prioridad se respeta el orden de llegada vía seq).
function enqueue(fn, priority = PRIORITY_NORMAL) {
  return new Promise((resolve, reject) => {
    const item = { fn, priority, seq: _seqCounter++, resolve, reject };
    // Insertar respetando prioridad (mayor primero), luego FIFO por seq.
    let i = _queue.length;
    while (i > 0 && _queue[i - 1].priority < priority) i--;
    _queue.splice(i, 0, item);
    _queueDepth = _queue.length;
    _procesarCola();
  });
}

// Worker: procesa un envío a la vez, con SEND_DELAY_MS entre uno y otro.
async function _procesarCola() {
  if (_draining) return;           // ya hay un envío en curso
  if (_queue.length === 0) return;
  _draining = true;

  const item = _queue.shift();
  _queueDepth = _queue.length;

  try {
    const result = await item.fn();
    item.resolve(result);
  } catch (err) {
    // Manejo de errores de contexto: reiniciar cliente + drenar cola.
    if (_isUploadFailure(err)) {
      log("WARN", "upload failed detectado — sesion WA degradada, se reconectara en proximo health check");
      _lastUploadError = Date.now();
    } else if (_isContextoMuerto(err)) {
      log("ERROR", "Contexto WA muerto en envio — reiniciando cliente y drenando cola", { error: err.message.slice(0, 80) });
      isReady = false;
      _lastUploadError = Date.now();
      _notificarError("Contexto WA muerto", `sendMessage fallo por contexto de Chrome perdido (${err.message.slice(0, 120)}). Reiniciando y drenando cola de ${_queue.length} pendientes.`);
      item.reject(err);
      _drenarCola("contexto de Chrome perdido");
      _draining = false;
      setTimeout(() => _reiniciarCliente(), 3000);
      return; // no aplicar el delay de 5s: no hay nada en cola y el cliente se reinicia
    }
    item.reject(err);
  }

  // Delay anti-ban antes del siguiente envío.
  setTimeout(() => {
    _draining = false;
    _procesarCola();
  }, SEND_DELAY_MS);
}

// ============================================================
// REINICIO PREVENTIVO PERIÓDICO — cada 6 horas reinicia Chrome
// para liberar RAM acumulada. Con 5 proyectos (SSFF horario,
// AVANCE, PBI, VPN, tablas) la RAM crece más rápido que antes,
// por eso ya no basta un único reinicio nocturno. Solo actúa
// si la cola está vacía (no interrumpe envíos en curso).
// ============================================================
const _REINICIO_PREVENTIVO_MS = 6 * 60 * 60 * 1000; // 6 horas
let _ultimoReinicioPreventivo = Date.now();

function _reinicioPreventivo() {
  if (_reconectando) return;
  const ahora = Date.now();
  if ((ahora - _ultimoReinicioPreventivo) < _REINICIO_PREVENTIVO_MS) return;
  if (_queueDepth > 0) {
    // No reprogramar el temporizador: reintentará en el próximo tick cuando la cola se vacíe.
    log("INFO", "Reinicio preventivo postergado — cola no vacia", { pendientes: _queueDepth });
    return;
  }
  _ultimoReinicioPreventivo = ahora;
  log("INFO", "Reinicio preventivo (cada 6h) — liberando RAM de Chrome acumulada");
  _reiniciarCliente();
}

// Verificar cada minuto si toca el reinicio preventivo
setInterval(_reinicioPreventivo, 60 * 1000);

// ============================================================
// SERVIDOR EXPRESS
// ============================================================
const app = express();
app.use(cors());
app.use(express.json({ limit: "50mb" }));

// Resuelve la prioridad del envío a partir del body.
// Acepta priority numérico directo, o el atajo priority:"avance".
// Por defecto PRIORITY_NORMAL (cortes horarios, capturas periódicas).
function _resolverPrioridad(body) {
  const p = body && body.priority;
  if (typeof p === "number") return p;
  if (typeof p === "string" && p.toLowerCase() === "avance") return PRIORITY_AVANCE;
  return PRIORITY_NORMAL;
}

// sendMessage() de whatsapp-web.js puede devolver `undefined` sin lanzar
// excepción cuando WhatsApp Web cambió algo en su protocolo (mismo síntoma
// que detecta el send-test periódico, ver más abajo). Antes esto quedaba
// completamente silencioso: el endpoint respondía success:true con
// messageId:null y no se generaba ningún log — indistinguible de un envío
// real. Se registra aquí como WARN y alimenta el mismo contador de fallos
// consecutivos que usa el send-test, así un patrón repetido en envíos
// reales (no solo en la prueba cada 3h) también dispara el reinicio.
function _registrarResultadoEnvio(kind, result, ctx) {
  if (result && result.id) {
    if (_sendTestFails > 0) log("INFO", "Envio recuperado tras fallos previos de protocolo");
    _sendTestFails = 0;
    log("INFO", kind, ctx);
    return;
  }
  _sendTestFails++;
  log("WARN", `${kind}: sendMessage devolvio resultado sin id (fallo #${_sendTestFails})`, ctx);
  if (_sendTestFails >= 2) {
    log("ERROR", "2 fallos consecutivos de protocolo en envios reales — protocolo WA degradado, reiniciando cliente");
    _notificarError(
      "Protocolo WA degradado (envio real)",
      `sendMessage devolvio resultado sin id en 2 envios consecutivos (ultimo: ${kind}). Es probable un cambio de protocolo de Meta. Reiniciando cliente automaticamente.`
    );
    _sendTestFails = 0;
    isReady = false;
    _reiniciarCliente();
  }
}

// Guard reutilizable: rechaza si el servidor no está listo o la cola está llena.
function _guardReady(res) {
  if (!isReady) {
    res.status(503).json({ error: "WhatsApp no esta listo" });
    return false;
  }
  if (queueFull()) {
    log("WARN", "Cola llena — rechazando envio para proteger RAM de Chrome", { depth: _queueDepth });
    res.status(503).json({ error: `Cola llena (${_queueDepth}/${QUEUE_MAX} mensajes pendientes). Reintenta en unos segundos.` });
    return false;
  }
  return true;
}

// --- Health check ---
// Además de isReady, verifica que el frame de Chrome siga respondiendo.
// Si hubo un "upload failed" reciente (último minuto), reporta degraded y fuerza reconexión.
app.get("/health", async (req, res) => {
  if (!isReady) {
    return res.json({ status: "not_ready", timestamp: new Date().toISOString() });
  }

  // Detectar upload failures recientes (en los últimos 90s)
  const ahoraMs = Date.now();
  if (_lastUploadError && (ahoraMs - _lastUploadError) < 90_000) {
    log("WARN", "Health: upload failure reciente detectado — marcando not_ready y reconectando");
    isReady = false;
    _lastUploadError = null;
    _reiniciarCliente();
    return res.json({ status: "not_ready", reason: "upload_failure", timestamp: new Date().toISOString() });
  }

  // Verificar que el frame de Chrome siga vivo
  try {
    await waClient.pupPage.evaluate(() => true);
    res.json({ status: "ready", timestamp: new Date().toISOString() });
  } catch (e) {
    log("WARN", "Health: Chrome no responde — marcando not_ready", { error: e.message });
    isReady = false;
    _reiniciarCliente();
    res.json({ status: "not_ready", reason: "chrome_unresponsive", timestamp: new Date().toISOString() });
  }
});

// --- Listar grupos ---
app.get("/list-groups", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no esta listo" });
  try {
    const chats = await waClient.getChats();
    const groups = chats
      .filter((c) => c.isGroup)
      .map((g) => ({ name: g.name, id: g.id._serialized, participants: g.participants?.length || 0 }));
    log("INFO", `Listados ${groups.length} grupos`);
    res.json(groups);
  } catch (err) {
    log("ERROR", "Error al listar grupos", { error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Buscar contactos por nombre ---
app.get("/list-contacts", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no esta listo" });
  const { name } = req.query;
  if (!name) return res.status(400).json({ error: "Parametro 'name' requerido" });
  try {
    const contacts = await waClient.getContacts();
    const filtered = contacts
      .filter((c) =>
        (c.name && c.name.toLowerCase().includes(name.toLowerCase())) ||
        (c.pushname && c.pushname.toLowerCase().includes(name.toLowerCase()))
      )
      .map((c) => ({
        name: c.name || c.pushname || "Sin nombre",
        pushname: c.pushname || "",
        id: c.id._serialized,
        number: c.id.user,
      }));
    log("INFO", `Encontrados ${filtered.length} contactos para "${name}"`);
    res.json(filtered);
  } catch (err) {
    log("ERROR", "Error al buscar contactos", { error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar texto ---
app.post("/send-text", async (req, res) => {
  if (!_guardReady(res)) return;
  const { to, message } = req.body;
  if (!to || !message) return res.status(400).json({ error: "Campos 'to' y 'message' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  try {
    const result = await enqueue(() => waClient.sendMessage(chatId, message), _resolverPrioridad(req.body));
    _registrarResultadoEnvio("Texto enviado", result, { to, chatId });
    res.json({ success: true, messageId: result?.id?._serialized || null });
  } catch (err) {
    log("ERROR", "Error al enviar texto", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar imagen (desde ruta local) ---
app.post("/send-image", async (req, res) => {
  if (!_guardReady(res)) return;
  const { to, image_path, caption } = req.body;
  if (!to || !image_path) return res.status(400).json({ error: "Campos 'to' y 'image_path' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  if (!fs.existsSync(image_path)) return res.status(404).json({ error: `Archivo no encontrado: ${image_path}` });
  try {
    const media = MessageMedia.fromFilePath(image_path);
    const result = await enqueue(() => waClient.sendMessage(chatId, media, { caption: caption || "" }), _resolverPrioridad(req.body));
    _registrarResultadoEnvio("Imagen enviada", result, { to, chatId, image_path });
    res.json({ success: true, messageId: result?.id?._serialized || null });
  } catch (err) {
    log("ERROR", "Error al enviar imagen", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar archivo (desde ruta local) ---
app.post("/send-file", async (req, res) => {
  if (!_guardReady(res)) return;
  const { to, file_path, caption } = req.body;
  if (!to || !file_path) return res.status(400).json({ error: "Campos 'to' y 'file_path' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  if (!fs.existsSync(file_path)) return res.status(404).json({ error: `Archivo no encontrado: ${file_path}` });
  try {
    const media = MessageMedia.fromFilePath(file_path);
    const result = await enqueue(() => waClient.sendMessage(chatId, media, { caption: caption || "" }), _resolverPrioridad(req.body));
    _registrarResultadoEnvio("Archivo enviado", result, { to, chatId, file_path });
    res.json({ success: true, messageId: result?.id?._serialized || null });
  } catch (err) {
    log("ERROR", "Error al enviar archivo", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// Timeout genérico: si `promise` no resuelve en `ms`, rechaza con un error
// cuyo mensaje contiene `label` — permite reusar la detección de contexto
// muerto (_isContextoMuerto busca substrings de mensaje) para cuelgues
// silenciosos que nunca lanzan Promise was collected / Protocol error.
function _withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, rej) => setTimeout(() => rej(new Error(`${label} timeout tras ${ms}ms — contexto de Chrome probablemente muerto`)), ms)),
  ]);
}

function _isMentionTimeout(err) {
  return err && err.message && err.message.includes("contexto de Chrome probablemente muerto");
}

// --- Enviar texto con menciones ---
app.post("/send-mention", async (req, res) => {
  if (!_guardReady(res)) return;
  const { to, message, mentions } = req.body;
  if (!to || !message) return res.status(400).json({ error: "Campos 'to' y 'message' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  try {
    const mentionIds = mentions || [];
    log("INFO", "Intentando enviar menciones", { to, chatId, mentionIds });

    // Resolver IDs a objetos Contact reales (necesario para que WA renderice menciones).
    // Cada resolución va con timeout de 20s: si getContactById() cuelga por un
    // contexto de Chrome muerto que no lanza excepción propia, no debe bloquear
    // el request (y con él, indirectamente, el resto del servidor) para siempre.
    const mentionContacts = [];
    for (const id of mentionIds) {
      try {
        const contact = await _withTimeout(waClient.getContactById(id), 20000, "getContactById");
        mentionContacts.push(contact);
      } catch (e) {
        if (_isMentionTimeout(e) || _isContextoMuerto(e)) {
          // Cuelgue real de Chrome: no tiene sentido seguir resolviendo el resto
          // de menciones ni intentar el sendMessage — mismo tratamiento que un
          // contexto muerto detectado en la cola (reiniciar cliente + drenar).
          log("ERROR", "getContactById colgado — contexto WA muerto, reiniciando cliente y drenando cola", { to, id, error: e.message });
          isReady = false;
          _lastUploadError = Date.now();
          _notificarError("Contexto WA muerto (send-mention)", `getContactById no respondio en 20s resolviendo menciones para "${to}". Reiniciando y drenando cola de ${_queue.length} pendientes.`);
          _drenarCola("contexto de Chrome perdido (timeout en send-mention)");
          setTimeout(() => _reiniciarCliente(), 3000);
          return res.status(503).json({ error: "Contexto de WhatsApp perdido resolviendo menciones. El servidor se esta reiniciando, reintenta en unos segundos." });
        }
        // Si no se puede resolver, construir objeto mínimo que whatsapp-web.js acepta
        log("WARN", `No se pudo resolver contacto ${id}, usando fallback`, { error: e.message });
        const [user] = id.split("@");
        mentionContacts.push({ id: { _serialized: id, user, server: "c.us" } });
      }
    }

    const _prio = _resolverPrioridad(req.body);
    let result;
    try {
      result = await enqueue(() =>
        waClient.sendMessage(chatId, message, { mentions: mentionContacts })
      , _prio);
    } catch (mentionErr) {
      log("WARN", "Menciones fallaron, enviando sin ellas", { to, error: mentionErr.message });
      result = await enqueue(() => waClient.sendMessage(chatId, message), _prio);
    }
    _registrarResultadoEnvio("Mensaje con menciones enviado", result, { to, chatId, menciones: mentionContacts.length });
    res.json({ success: true, messageId: result?.id?._serialized || null });
  } catch (err) {
    log("ERROR", "Error al enviar menciones", { to, error: err.message, stack: err.stack });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar link ---
app.post("/send-link", async (req, res) => {
  if (!_guardReady(res)) return;
  const { to, url, description } = req.body;
  if (!to || !url) return res.status(400).json({ error: "Campos 'to' y 'url' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  try {
    const message = description ? `${description}\n${url}` : url;
    const result = await enqueue(() => waClient.sendMessage(chatId, message), _resolverPrioridad(req.body));
    _registrarResultadoEnvio("Link enviado", result, { to, chatId, url });
    res.json({ success: true, messageId: result?.id?._serialized || null });
  } catch (err) {
    log("ERROR", "Error al enviar link", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// ============================================================
// SEND-TEST PERIÓDICO — detecta fallos de protocolo antes de que
// fallen envíos reales. Cada 3 horas envía un mensaje silencioso
// al propio número (MY_WA_NUMBER en .env). Si sendMessage devuelve
// undefined dos veces seguidas, dispara alerta y reinicia el cliente.
// ============================================================
const _SEND_TEST_INTERVAL_MS = 3 * 60 * 60 * 1000; // 3 horas
let _sendTestFails = 0;

async function _runSendTest() {
  if (!isReady || _reconectando) return;

  // Usa OWNER_WA_ID (ya definido en .env para alertas internas)
  const myNumber = process.env.OWNER_WA_ID;
  if (!myNumber) return; // variable no configurada → omitir silenciosamente

  const chatId = myNumber.includes("@") ? myNumber : `${myNumber}@c.us`;
  try {
    const result = await waClient.sendMessage(chatId, "✔️ [auto-test] servidor WA operativo");
    if (!result || !result.id) {
      _sendTestFails++;
      log("WARN", `Send-test: sendMessage devolvio undefined (fallo #${_sendTestFails})`, { chatId });
      if (_sendTestFails >= 2) {
        log("ERROR", "Send-test: 2 fallos consecutivos — protocolo WA degradado, reiniciando cliente");
        _notificarError(
          "Protocolo WA degradado (send-test)",
          `sendMessage devolvio undefined en 2 pruebas consecutivas. Es probable un cambio de protocolo de Meta. Reiniciando cliente automaticamente.`
        );
        _sendTestFails = 0;
        isReady = false;
        _reiniciarCliente();
      }
    } else {
      if (_sendTestFails > 0) log("INFO", "Send-test: recuperado tras fallos anteriores");
      _sendTestFails = 0;
      log("INFO", "Send-test OK", { chatId });
    }
  } catch (e) {
    _sendTestFails++;
    log("WARN", `Send-test: error al enviar (fallo #${_sendTestFails})`, { error: e.message });
    if (_sendTestFails >= 2) {
      log("ERROR", "Send-test: 2 errores consecutivos — reiniciando cliente");
      _notificarError("Protocolo WA degradado (send-test)", `Error en send-test: ${e.message}`);
      _sendTestFails = 0;
      isReady = false;
      _reiniciarCliente();
    }
  }
}

// Primera prueba 5 minutos después de arrancar (esperar a que el cliente esté listo)
setTimeout(() => {
  _runSendTest();
  setInterval(_runSendTest, _SEND_TEST_INTERVAL_MS);
}, 5 * 60 * 1000);

// ============================================================
// AUTO-UPDATE SEMANAL — actualiza whatsapp-web.js desde GitHub
// (el repo tiene fixes antes que npm). Se ejecuta los domingos
// a las 3 AM hora local para minimizar impacto operativo.
// ============================================================
function _autoUpdateWwjs() {
  const ahora = new Date();
  // Solo ejecutar domingos (0) entre 03:00 y 03:59
  if (ahora.getDay() !== 0 || ahora.getHours() !== 3) return;

  log("INFO", "Auto-update: iniciando actualización de whatsapp-web.js desde GitHub...");
  const { exec } = require("child_process");
  const cmd = "npm install github:pedroslopez/whatsapp-web.js --save";
  exec(cmd, { cwd: __dirname }, (err, stdout, stderr) => {
    if (err) {
      log("WARN", "Auto-update: error al actualizar whatsapp-web.js", { error: err.message, stderr });
      return;
    }
    // Verificar si cambió algo comparando el gitHead del package instalado
    try {
      const pkgPath = path.join(__dirname, "node_modules", "whatsapp-web.js", "package.json");
      const pkg = JSON.parse(fs.readFileSync(pkgPath, "utf8"));
      log("INFO", "Auto-update: whatsapp-web.js actualizado", {
        version: pkg.version,
        gitHead: pkg.gitHead || "N/A",
      });
    } catch (_) {
      log("INFO", "Auto-update: actualización completada", { stdout: stdout.trim() });
    }
    // Reiniciar el cliente para que tome la nueva versión
    log("INFO", "Auto-update: reiniciando cliente para aplicar nueva version...");
    setTimeout(() => _reiniciarCliente(), 3000);
  });
}

// Verificar cada hora si es momento de hacer el update
setInterval(_autoUpdateWwjs, 60 * 60 * 1000);

// ============================================================
// INICIAR TODO — con guard de INSTANCIA ÚNICA
// Si el puerto 8002 ya está en uso, significa que otro wa_server.js
// ya está corriendo (watchdog + reinicio interno pisándose). En ese
// caso NO arrancamos Chrome: abortamos limpio para no tener dos
// clientes WhatsApp compitiendo por la misma sesión (causa de
// congelamientos y "a veces no inicia" observados el 17/07).
// ============================================================
const server = app.listen(PORT);

server.on("listening", () => {
  log("INFO", `Servidor HTTP escuchando en puerto ${PORT}`);
  startWhatsApp();
});

server.on("error", (err) => {
  if (err.code === "EADDRINUSE") {
    log("ERROR", `Puerto ${PORT} ya en uso — ya existe otra instancia de wa_server.js. Abortando esta instancia para evitar sesiones WA duplicadas.`);
    process.exit(1);
  } else {
    log("ERROR", "Error al iniciar servidor HTTP", { error: err.message });
    process.exit(1);
  }
});
