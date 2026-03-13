const path = require("path");
const os = require("os");
const { mkdtemp, readFile, rm, writeFile } = require("fs/promises");
const { pathToFileURL } = require("url");
const { app, BrowserWindow, clipboard, dialog, ipcMain, screen, shell } = require("electron");
const { ensureMboxDatabase, searchMessages, loadMessageById, getMessageDateBounds } = require("./src/mboxStore");
const { isPstFilePath, ensurePstConvertedToMbox } = require("./src/pstConverter");

const DEFAULT_PAGE_SIZE = 200;
const OPEN_PROGRESS_EVENT = "mbox-index-progress";
const PREVIEW_WINDOW_DEFAULT_BOUNDS = { width: 960, height: 760 };
const PREVIEW_WINDOW_MIN_BOUNDS = { width: 480, height: 360 };
let attachmentPreviewWindow = null;
let attachmentPreviewWindowState = null;

function createWindow() {
  const window = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 980,
    minHeight: 640,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  window.webContents.setWindowOpenHandler(({ url }) => {
    if (isInternalAppUrl(url)) {
      return { action: "allow" };
    }
    return { action: "deny" };
  });

  window.webContents.on("will-navigate", (event, url) => {
    if (isInternalAppUrl(url)) {
      return;
    }
    event.preventDefault();
  });

  window.loadFile(path.join(__dirname, "src/renderer/index.html"));
}

ipcMain.handle("open-mbox", async (event) => {
  const result = await dialog.showOpenDialog({
    title: "Open mailbox file",
    properties: ["openFile"],
    filters: [
      { name: "Mailbox Files", extensions: ["mbox", "pst"] },
      { name: "Mbox Files", extensions: ["mbox"] },
      { name: "Outlook PST Files", extensions: ["pst"] },
      { name: "All Files", extensions: ["*"] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true };
  }

  const filePath = result.filePaths[0];
  const openedAsPst = isPstFilePath(filePath);
  let sourcePath = filePath;
  let pstConversion = null;

  emitOpenProgress(event.sender, {
    phase: "preparing",
    filePath,
    totalBytes: 0,
    bytesRead: 0,
    messagesIndexed: 0
  });

  if (openedAsPst) {
    pstConversion = await ensurePstConvertedToMbox(filePath, {
      onProgress: (payload) => {
        emitOpenProgress(event.sender, {
          filePath,
          ...payload
        });
      }
    });
    sourcePath = pstConversion.mboxPath;
  }

  const indexing = await ensureMboxDatabase(sourcePath, event.sender);
  const firstPage = searchMessages(indexing.dbPath, "", DEFAULT_PAGE_SIZE, 0);
  const dateBounds = getMessageDateBounds(indexing.dbPath);

  return {
    canceled: false,
    filePath,
    sourcePath,
    sourceType: openedAsPst ? "pst" : "mbox",
    pstConversion,
    dbPath: indexing.dbPath,
    total: indexing.totalMessages,
    messages: firstPage.messages,
    offset: firstPage.offset,
    limit: firstPage.limit,
    resultTotal: firstPage.total,
    dateRange: dateBounds
      ? {
          from: dateBounds.minDateTs,
          to: dateBounds.maxDateTs,
          count: dateBounds.datedCount
        }
      : null
  };
});

ipcMain.handle("search-messages", async (_, payload) => {
  const dbPath = typeof payload?.dbPath === "string" ? payload.dbPath : "";
  const query = typeof payload?.query === "string" ? payload.query : "";
  const limit = payload?.limit;
  const offset = payload?.offset;
  const dateFrom = payload?.dateFrom;
  const dateTo = payload?.dateTo;
  const senderQuery = typeof payload?.senderQuery === "string" ? payload.senderQuery : "";
  const recipientQuery = typeof payload?.recipientQuery === "string" ? payload.recipientQuery : "";
  const subjectQuery = typeof payload?.subjectQuery === "string" ? payload.subjectQuery : "";
  const attachmentsOnly = Boolean(payload?.attachmentsOnly);

  if (!dbPath) {
    return { total: 0, offset: 0, limit: 0, messages: [] };
  }

  return searchMessages(dbPath, query, limit, offset, {
    dateFrom,
    dateTo,
    senderQuery,
    recipientQuery,
    subjectQuery,
    attachmentsOnly
  });
});

ipcMain.handle("get-message", async (_, payload) => {
  const dbPath = typeof payload?.dbPath === "string" ? payload.dbPath : "";
  const id = payload?.id;

  if (!dbPath || id === undefined || id === null) {
    return null;
  }

  return loadMessageById(dbPath, id);
});

ipcMain.handle("save-attachment", async (_, payload) => {
  const fileName = typeof payload?.fileName === "string" ? payload.fileName : "attachment.bin";
  const base64 = typeof payload?.base64 === "string" ? payload.base64 : "";

  if (!base64) {
    return { canceled: true, error: "No attachment data" };
  }

  const result = await dialog.showSaveDialog({
    title: "Save attachment",
    defaultPath: fileName
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  const data = Buffer.from(base64, "base64");
  await writeFile(result.filePath, data);
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle("save-message-eml", async (_, payload) => {
  const rawFileName = typeof payload?.fileName === "string" ? payload.fileName : "message.eml";
  const fileName = normalizeEmlFileName(rawFileName);
  const emlSource = typeof payload?.emlSource === "string" ? payload.emlSource : "";

  if (!emlSource.trim()) {
    return { canceled: true, error: "No message source available" };
  }

  const result = await dialog.showSaveDialog({
    title: "Save message as EML",
    defaultPath: fileName
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  const normalized = emlSource.replace(/\r?\n/g, "\r\n");
  await writeFile(result.filePath, normalized, "utf8");
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle("open-external", async (_, payload) => {
  const rawUrl = typeof payload?.url === "string" ? payload.url : "";
  const opened = await openExternalUrl(rawUrl);
  return { opened };
});

ipcMain.handle("copy-to-clipboard", async (_, payload) => {
  const text = typeof payload?.text === "string" ? payload.text : "";
  if (!text) {
    return { copied: false };
  }

  try {
    clipboard.writeText(text);
    return { copied: true };
  } catch (error) {
    console.error("Failed to copy text to clipboard.", error);
    return { copied: false };
  }
});

ipcMain.handle("open-attachment-preview", async (event, payload) => {
  const fileName = typeof payload?.fileName === "string" ? payload.fileName : "attachment";
  const contentType = typeof payload?.contentType === "string" ? payload.contentType : "";
  const base64 = typeof payload?.base64 === "string" ? payload.base64 : "";
  let tempDir = "";

  if (!base64 || !isPreviewableContentType(contentType)) {
    return { opened: false };
  }

  try {
    tempDir = await mkdtemp(path.join(os.tmpdir(), "mbox-viewer-preview-"));
    const tempFilePath = path.join(tempDir, buildPreviewFileName(fileName, contentType));
    await writeFile(tempFilePath, Buffer.from(base64, "base64"));

    const parentWindow = BrowserWindow.fromWebContents(event.sender) || null;
    const previewWindow = await createAttachmentPreviewWindow(parentWindow);
    const previewUrl = buildAttachmentPreviewPageUrl(tempFilePath, fileName, contentType);

    await previewWindow.loadURL(previewUrl);
    previewWindow.setTitle(fileName || "Attachment Preview");
    previewWindow.show();
    previewWindow.focus();

    cleanupAttachmentPreviewPath(previewWindow.__previewTempDir);
    previewWindow.__previewTempDir = tempDir;

    return { opened: true };
  } catch (error) {
    cleanupAttachmentPreviewPath(tempDir);
    console.error("Failed to open attachment preview.", error);
    return { opened: false };
  }
});

function normalizeEmlFileName(input) {
  const stripped = String(input || "message")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  const base = stripped || "message";
  return base.toLowerCase().endsWith(".eml") ? base : `${base}.eml`;
}

function isInternalAppUrl(url) {
  const value = String(url || "").trim();
  return value.startsWith("file:") || value.startsWith("devtools:");
}

function isPreviewableContentType(contentType) {
  const value = String(contentType || "").toLowerCase();
  return value.startsWith("image/") || value === "application/pdf";
}

function buildPreviewFileName(fileName, contentType) {
  const safeBaseName = String(fileName || "attachment")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .trim();

  if (path.extname(safeBaseName)) {
    return safeBaseName;
  }

  if (contentType === "application/pdf") {
    return `${safeBaseName || "attachment"}.pdf`;
  }

  const imageSubtype = String(contentType || "")
    .toLowerCase()
    .match(/^image\/([a-z0-9.+-]+)$/)?.[1];

  return `${safeBaseName || "attachment"}.${imageSubtype || "bin"}`;
}

function buildAttachmentPreviewPageUrl(filePath, fileName, contentType) {
  const previewPageUrl = pathToFileURL(path.join(__dirname, "src/renderer/attachmentPreview.html"));
  previewPageUrl.searchParams.set("file", pathToFileURL(filePath).toString());
  previewPageUrl.searchParams.set("name", fileName || "Attachment Preview");
  previewPageUrl.searchParams.set("type", contentType || "application/octet-stream");
  return previewPageUrl.toString();
}

async function createAttachmentPreviewWindow(parentWindow) {
  if (attachmentPreviewWindow && !attachmentPreviewWindow.isDestroyed()) {
    return attachmentPreviewWindow;
  }

  const restoredBounds = await loadAttachmentPreviewWindowState();
  const previewWindow = new BrowserWindow({
    width: restoredBounds?.width || PREVIEW_WINDOW_DEFAULT_BOUNDS.width,
    height: restoredBounds?.height || PREVIEW_WINDOW_DEFAULT_BOUNDS.height,
    x: Number.isFinite(restoredBounds?.x) ? restoredBounds.x : undefined,
    y: Number.isFinite(restoredBounds?.y) ? restoredBounds.y : undefined,
    minWidth: PREVIEW_WINDOW_MIN_BOUNDS.width,
    minHeight: PREVIEW_WINDOW_MIN_BOUNDS.height,
    resizable: true,
    autoHideMenuBar: true,
    show: false,
    parent: parentWindow || undefined,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  previewWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  previewWindow.webContents.on("will-navigate", (event, url) => {
    if (isInternalAppUrl(url)) {
      return;
    }
    event.preventDefault();
  });

  previewWindow.on("close", () => {
    saveAttachmentPreviewWindowState(previewWindow);
  });
  previewWindow.on("closed", () => {
    cleanupAttachmentPreviewPath(previewWindow.__previewTempDir);
    attachmentPreviewWindow = null;
  });

  if (restoredBounds?.isMaximized) {
    previewWindow.once("ready-to-show", () => {
      if (!previewWindow.isDestroyed()) {
        previewWindow.maximize();
      }
    });
  }

  attachmentPreviewWindow = previewWindow;
  return previewWindow;
}

function cleanupAttachmentPreviewPath(targetPath) {
  if (!targetPath) {
    return;
  }

  rm(targetPath, { recursive: true, force: true }).catch((error) => {
    console.error(`Failed to clean attachment preview path: ${targetPath}`, error);
  });
}

async function loadAttachmentPreviewWindowState() {
  if (attachmentPreviewWindowState !== null) {
    return attachmentPreviewWindowState;
  }

  try {
    const raw = await readFile(getAttachmentPreviewWindowStatePath(), "utf8");
    attachmentPreviewWindowState = sanitizeAttachmentPreviewWindowState(JSON.parse(raw));
  } catch {
    attachmentPreviewWindowState = null;
  }

  return attachmentPreviewWindowState;
}

function saveAttachmentPreviewWindowState(window) {
  if (!window || window.isDestroyed()) {
    return;
  }

  const normalBounds = window.getNormalBounds();
  const nextState = sanitizeAttachmentPreviewWindowState({
    ...normalBounds,
    isMaximized: window.isMaximized()
  });

  attachmentPreviewWindowState = nextState;
  writeFile(getAttachmentPreviewWindowStatePath(), JSON.stringify(nextState), "utf8").catch((error) => {
    console.error("Failed to save attachment preview window state.", error);
  });
}

function sanitizeAttachmentPreviewWindowState(value) {
  const width = clampInteger(value?.width, PREVIEW_WINDOW_MIN_BOUNDS.width, 3200, PREVIEW_WINDOW_DEFAULT_BOUNDS.width);
  const height = clampInteger(value?.height, PREVIEW_WINDOW_MIN_BOUNDS.height, 2400, PREVIEW_WINDOW_DEFAULT_BOUNDS.height);
  const result = {
    width,
    height,
    isMaximized: Boolean(value?.isMaximized)
  };

  if (Number.isFinite(value?.x) && Number.isFinite(value?.y)) {
    const candidate = {
      x: Math.round(value.x),
      y: Math.round(value.y),
      width,
      height
    };

    if (isRectangleVisibleOnAnyDisplay(candidate)) {
      result.x = candidate.x;
      result.y = candidate.y;
    }
  }

  return result;
}

function isRectangleVisibleOnAnyDisplay(rect) {
  const displays = screen.getAllDisplays();
  return displays.some((display) => getIntersectionArea(rect, display.workArea) > 0);
}

function getIntersectionArea(a, b) {
  const left = Math.max(a.x, b.x);
  const top = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.width, b.x + b.width);
  const bottom = Math.min(a.y + a.height, b.y + b.height);

  if (right <= left || bottom <= top) {
    return 0;
  }

  return (right - left) * (bottom - top);
}

function clampInteger(value, min, max, fallback) {
  const parsed = Math.round(Number(value));
  if (!Number.isFinite(parsed)) {
    return fallback;
  }
  if (parsed < min) {
    return min;
  }
  if (parsed > max) {
    return max;
  }
  return parsed;
}

function getAttachmentPreviewWindowStatePath() {
  return path.join(app.getPath("userData"), "attachment-preview-window.json");
}

function normalizeExternalUrl(rawUrl) {
  const value = String(rawUrl || "").trim();
  if (!value) {
    return "";
  }

  try {
    const parsed = new URL(value);
    const protocol = parsed.protocol.toLowerCase();
    if (!["http:", "https:", "mailto:", "tel:"].includes(protocol)) {
      return "";
    }
    return parsed.toString();
  } catch {
    return "";
  }
}

async function openExternalUrl(rawUrl) {
  const url = normalizeExternalUrl(rawUrl);
  if (!url) {
    return false;
  }

  try {
    await shell.openExternal(url);
    return true;
  } catch (error) {
    console.error(`Failed to open external url: ${url}`, error);
    return false;
  }
}

function emitOpenProgress(sender, payload) {
  if (!sender || sender.isDestroyed()) {
    return;
  }
  sender.send(OPEN_PROGRESS_EVENT, payload);
}

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
