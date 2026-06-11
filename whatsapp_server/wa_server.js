const { Client, LocalAuth, MessageMedia } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");
const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

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

function startWhatsApp() {
  log("INFO", "Iniciando cliente WhatsApp...");

  waClient = new Client({
    authStrategy: new LocalAuth({
      clientId: "automation_session",
      dataPath: SESSION_DIR,
    }),
    puppeteer: {
      headless: true,
      executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-extensions",
        "--disable-background-networking",
        "--disable-default-apps",
        "--no-first-run",
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding",
        "--disable-features=TranslateUI",
      ],
    },
  });

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
    // Force-kill cualquier proceso Chrome huerfano antes de reiniciar
    try {
      execSync("taskkill /F /IM chrome.exe /T 2>nul", { stdio: "ignore" });
      log("INFO", "Procesos Chrome terminados forzosamente");
    } catch (_) { /* ignorar si no hay Chrome */ }
  }
  waClient = null;
  isReady  = false;
  log("INFO", "Relanzando cliente WhatsApp...");
  _reconectando = false;
  startWhatsApp();
}

// Watchdog: cada 5 minutos verifica que Chrome siga vivo haciendo un ping real
// Si detecta que el frame está caído, reinicia antes de que llegue un envío real
setInterval(async () => {
  if (!isReady || _reconectando) return;
  try {
    await waClient.pupPage.evaluate(() => true);
  } catch (e) {
    if (_isDetachedFrame(e)) {
      log("ERROR", "Watchdog: Chrome caido (detached Frame) — reiniciando cliente");
      isReady = false;
      _reiniciarCliente();
    }
  }
}, 5 * 60 * 1000);

// ============================================================
// COLA DE ENVÍO — mínimo 5 s entre mensajes salientes
// ============================================================
const SEND_DELAY_MS = 5000;
let sendQueue = Promise.resolve();

function _isDetachedFrame(err) {
  return err && err.message && err.message.includes("detached Frame");
}

function enqueue(fn) {
  sendQueue = sendQueue.then(
    () => fn().then(
      (result) => new Promise((res) => setTimeout(() => res(result), SEND_DELAY_MS)),
      (err) => new Promise((_, rej) => {
        if (_isDetachedFrame(err)) {
          log("ERROR", "Chrome desconectado (detached Frame) — reiniciando cliente en 5s");
          isReady = false;
          setTimeout(() => _reiniciarCliente(), 5000);
        }
        setTimeout(() => rej(err), SEND_DELAY_MS);
      })
    )
  );
  return sendQueue;
}

// ============================================================
// SERVIDOR EXPRESS
// ============================================================
const app = express();
app.use(cors());
app.use(express.json({ limit: "50mb" }));

// --- Health check ---
app.get("/health", (req, res) => {
  res.json({ status: isReady ? "ready" : "not_ready", timestamp: new Date().toISOString() });
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
  if (!isReady) return res.status(503).json({ error: "WhatsApp no esta listo" });
  const { to, message } = req.body;
  if (!to || !message) return res.status(400).json({ error: "Campos 'to' y 'message' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  try {
    const result = await enqueue(() => waClient.sendMessage(chatId, message));
    log("INFO", "Texto enviado", { to, chatId });
    res.json({ success: true, messageId: result.id._serialized });
  } catch (err) {
    log("ERROR", "Error al enviar texto", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar imagen (desde ruta local) ---
app.post("/send-image", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no esta listo" });
  const { to, image_path, caption } = req.body;
  if (!to || !image_path) return res.status(400).json({ error: "Campos 'to' y 'image_path' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  if (!fs.existsSync(image_path)) return res.status(404).json({ error: `Archivo no encontrado: ${image_path}` });
  try {
    const media = MessageMedia.fromFilePath(image_path);
    const result = await enqueue(() => waClient.sendMessage(chatId, media, { caption: caption || "" }));
    log("INFO", "Imagen enviada", { to, chatId, image_path });
    res.json({ success: true, messageId: result.id._serialized });
  } catch (err) {
    log("ERROR", "Error al enviar imagen", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar archivo (desde ruta local) ---
app.post("/send-file", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no esta listo" });
  const { to, file_path, caption } = req.body;
  if (!to || !file_path) return res.status(400).json({ error: "Campos 'to' y 'file_path' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  if (!fs.existsSync(file_path)) return res.status(404).json({ error: `Archivo no encontrado: ${file_path}` });
  try {
    const media = MessageMedia.fromFilePath(file_path);
    const result = await enqueue(() => waClient.sendMessage(chatId, media, { caption: caption || "" }));
    log("INFO", "Archivo enviado", { to, chatId, file_path });
    res.json({ success: true, messageId: result.id._serialized });
  } catch (err) {
    log("ERROR", "Error al enviar archivo", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar texto con menciones ---
app.post("/send-mention", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no esta listo" });
  const { to, message, mentions } = req.body;
  if (!to || !message) return res.status(400).json({ error: "Campos 'to' y 'message' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  try {
    const mentionIds = mentions || [];
    log("INFO", "Intentando enviar menciones", { to, chatId, mentionIds });

    let result;
    try {
      result = await enqueue(() =>
        waClient.sendMessage(chatId, message, { mentions: mentionIds })
      );
    } catch (mentionErr) {
      // Si falla con menciones, enviar como texto plano sin ellas
      log("WARN", "Menciones fallaron, enviando sin ellas", { to, error: mentionErr.message });
      result = await enqueue(() => waClient.sendMessage(chatId, message));
    }
    log("INFO", "Mensaje con menciones enviado", { to, chatId, menciones: mentionIds.length });
    res.json({ success: true, messageId: result.id._serialized });
  } catch (err) {
    log("ERROR", "Error al enviar menciones", { to, error: err.message, stack: err.stack });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar link ---
app.post("/send-link", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no esta listo" });
  const { to, url, description } = req.body;
  if (!to || !url) return res.status(400).json({ error: "Campos 'to' y 'url' requeridos" });
  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });
  try {
    const message = description ? `${description}\n${url}` : url;
    const result = await enqueue(() => waClient.sendMessage(chatId, message));
    log("INFO", "Link enviado", { to, chatId, url });
    res.json({ success: true, messageId: result.id._serialized });
  } catch (err) {
    log("ERROR", "Error al enviar link", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// ============================================================
// INICIAR TODO
// ============================================================
app.listen(PORT, () => {
  log("INFO", `Servidor HTTP escuchando en puerto ${PORT}`);
});

startWhatsApp();
