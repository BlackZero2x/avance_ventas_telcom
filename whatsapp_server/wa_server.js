const { create, Client } = require("@open-wa/wa-automate");
const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");

// ============================================================
// CONFIGURACIÓN
// ============================================================
const PORT = 8002;
const SESSION_DIR = path.join(__dirname, "session_data");
const LOG_DIR = path.join(__dirname, "logs");
const CONFIG_PATH = path.join(__dirname, "config.json");

// Crear carpetas si no existen
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
  } catch (err) {
    log("WARN", "No se pudo cargar config.json, usando config vacía");
    return { groups: {}, contacts: {}, my_number: "" };
  }
}

// ============================================================
// RESOLVER DESTINATARIO
// ============================================================
function resolveRecipient(nameOrId) {
  const config = loadConfig();

  // Si ya es un ID de WhatsApp (contiene @), usarlo directo
  if (nameOrId.includes("@")) return nameOrId;

  // Buscar en grupos
  if (config.groups && config.groups[nameOrId]) return config.groups[nameOrId];

  // Buscar en contactos
  if (config.contacts && config.contacts[nameOrId]) return config.contacts[nameOrId];

  return null;
}

// ============================================================
// INICIAR CLIENTE WHATSAPP
// ============================================================
let waClient = null;
let isReady = false;

async function startWhatsApp() {
  log("INFO", "Iniciando cliente WhatsApp...");

  try {
    waClient = await create({
      sessionId: "automation_session",
      sessionDataPath: SESSION_DIR,
      headless: true,
      qrTimeout: 60,
      authTimeout: 120,
      cacheEnabled: false,
      useChrome: true,
      killProcessOnBrowserClose: false,
      throwErrorOnTosBlock: false,
      qrRefreshS: 15,
      logConsole: false,
      popup: false,
    });

    isReady = true;
    log("INFO", "✅ Cliente WhatsApp listo");

    // Manejar desconexión
    waClient.onStateChanged((state) => {
      log("INFO", "Estado WhatsApp cambiado", { state });
      if (state === "CONFLICT" || state === "UNLAUNCHED" || state === "UNPAIRED") {
        isReady = false;
        log("WARN", "Sesión desconectada. Reiniciar servidor para re-escanear QR.");
      }
    });

    return waClient;
  } catch (err) {
    log("ERROR", "Error al iniciar WhatsApp", { error: err.message });
    throw err;
  }
}

// ============================================================
// COLA DE ENVÍO — garantiza mínimo 5 s entre mensajes salientes
// ============================================================
const SEND_DELAY_MS = 5000;
let sendQueue = Promise.resolve();

function enqueue(fn) {
  sendQueue = sendQueue.then(() => fn()).then(
    (result) => { return new Promise((res) => setTimeout(() => res(result), SEND_DELAY_MS)); },
    (err)    => { return new Promise((_, rej) => setTimeout(() => rej(err),    SEND_DELAY_MS)); }
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
  res.json({
    status: isReady ? "ready" : "not_ready",
    timestamp: new Date().toISOString(),
  });
});

// --- Listar grupos ---
app.get("/list-groups", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no está listo" });

  try {
    const chats = await waClient.getAllGroups();
    const groups = chats.map((g) => ({
      name: g.name || g.formattedTitle || "Sin nombre",
      id: g.id,
      participants: g.groupMetadata ? g.groupMetadata.participants.length : 0,
    }));
    log("INFO", `Listados ${groups.length} grupos`);
    res.json(groups);
  } catch (err) {
    log("ERROR", "Error al listar grupos", { error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Buscar contactos por nombre ---
app.get("/list-contacts", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no está listo" });

  const { name } = req.query;
  if (!name) return res.status(400).json({ error: "Parámetro 'name' requerido" });

  try {
    const contacts = await waClient.getAllContacts();
    const filtered = contacts
      .filter(
        (c) =>
          c.name && c.name.toLowerCase().includes(name.toLowerCase()) ||
          c.pushname && c.pushname.toLowerCase().includes(name.toLowerCase())
      )
      .map((c) => ({
        name: c.name || c.pushname || "Sin nombre",
        pushname: c.pushname || "",
        id: c.id,
        number: c.id.replace("@c.us", ""),
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
  if (!isReady) return res.status(503).json({ error: "WhatsApp no está listo" });

  const { to, message } = req.body;
  if (!to || !message) return res.status(400).json({ error: "Campos 'to' y 'message' requeridos" });

  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });

  try {
    const result = await enqueue(() => waClient.sendText(chatId, message));
    log("INFO", "Texto enviado", { to, chatId });
    res.json({ success: true, messageId: result });
  } catch (err) {
    log("ERROR", "Error al enviar texto", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar imagen (desde ruta local) ---
app.post("/send-image", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no está listo" });

  const { to, image_path, caption } = req.body;
  if (!to || !image_path) return res.status(400).json({ error: "Campos 'to' y 'image_path' requeridos" });

  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });

  // Verificar que el archivo existe
  if (!fs.existsSync(image_path)) {
    return res.status(404).json({ error: `Archivo no encontrado: ${image_path}` });
  }

  try {
    // Convertir a base64 para enviar
    const imageData = fs.readFileSync(image_path);
    const base64 = `data:image/png;base64,${imageData.toString("base64")}`;

    const result = await enqueue(() => waClient.sendImage(chatId, base64, "screenshot.png", caption || ""));
    log("INFO", "Imagen enviada", { to, chatId, image_path });
    res.json({ success: true, messageId: result });
  } catch (err) {
    log("ERROR", "Error al enviar imagen", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar archivo (desde ruta local) ---
app.post("/send-file", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no está listo" });

  const { to, file_path, caption } = req.body;
  if (!to || !file_path) return res.status(400).json({ error: "Campos 'to' y 'file_path' requeridos" });

  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });

  if (!fs.existsSync(file_path)) {
    return res.status(404).json({ error: `Archivo no encontrado: ${file_path}` });
  }

  try {
    const filename = path.basename(file_path);
    const result = await enqueue(() => waClient.sendFile(chatId, file_path, filename, caption || ""));
    log("INFO", "Archivo enviado", { to, chatId, file_path });
    res.json({ success: true, messageId: result });
  } catch (err) {
    log("ERROR", "Error al enviar archivo", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar texto con menciones ---
app.post("/send-mention", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no está listo" });

  const { to, message, mentions } = req.body;
  if (!to || !message) return res.status(400).json({ error: "Campos 'to' y 'message' requeridos" });

  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });

  try {
    // mentions: array de IDs (ej. ["51962969371@c.us"]) o null para mencionar a todos los del texto
    const result = await enqueue(() => waClient.sendTextWithMentions(chatId, message, false, mentions || undefined));
    log("INFO", "Mensaje con menciones enviado", { to, chatId, mentions });
    res.json({ success: true, messageId: result });
  } catch (err) {
    log("ERROR", "Error al enviar menciones", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// --- Enviar link con mensaje ---
app.post("/send-link", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp no está listo" });

  const { to, url, description } = req.body;
  if (!to || !url) return res.status(400).json({ error: "Campos 'to' y 'url' requeridos" });

  const chatId = resolveRecipient(to);
  if (!chatId) return res.status(404).json({ error: `Destinatario "${to}" no encontrado en config` });

  try {
    const message = description ? `${description}\n${url}` : url;
    const result = await enqueue(() => waClient.sendText(chatId, message));
    log("INFO", "Link enviado", { to, chatId, url });
    res.json({ success: true, messageId: result });
  } catch (err) {
    log("ERROR", "Error al enviar link", { to, error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// ============================================================
// INICIAR TODO
// ============================================================
async function main() {
  app.listen(PORT, () => {
    log("INFO", `Servidor HTTP escuchando en puerto ${PORT}`);
  });

  await startWhatsApp();
}

main().catch((err) => {
  log("ERROR", "Error fatal al iniciar", { error: err.message });
  process.exit(1);
});
