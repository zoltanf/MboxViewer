const path = require("path");
const os = require("os");
const { mkdtemp, readFile, rm, stat, writeFile } = require("fs/promises");
const { pathToFileURL } = require("url");
const { app, BrowserWindow, clipboard, dialog, ipcMain, screen, shell } = require("electron");
const {
  ensureMboxDatabase,
  getReusableDatabaseInfo,
  searchMessages,
  loadMessageById,
  getMessageDateBounds
} = require("./src/mboxStore");
const { parseMessageChunk } = require("./src/mboxParser");
const { isPstFilePath, ensurePstConvertedToMbox } = require("./src/pstConverter");

const DEFAULT_PAGE_SIZE = 200;
const OPEN_PROGRESS_EVENT = "mbox-index-progress";
const OPEN_MAILBOX_REQUEST_EVENT = "open-mailbox-request";
const PREVIEW_WINDOW_DEFAULT_BOUNDS = { width: 960, height: 760 };
const PREVIEW_WINDOW_MIN_BOUNDS = { width: 480, height: 360 };
let attachmentPreviewWindow = null;
let attachmentPreviewWindowState = null;
let mainWindow = null;
let pendingOpenFilePath = "";

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

  window.on("closed", () => {
    if (mainWindow === window) {
      mainWindow = null;
    }
  });

  window.webContents.on("did-finish-load", () => {
    flushPendingMailboxOpenRequest(window);
  });

  window.loadFile(path.join(__dirname, "src/renderer/index.html"));
  mainWindow = window;
  return window;
}

ipcMain.handle("open-mbox", async (event) => {
  const result = await dialog.showOpenDialog({
    title: "Open email or mailbox file",
    properties: ["openFile"],
    filters: [
      { name: "Email Files", extensions: ["mbox", "pst", "eml"] },
      { name: "Mbox Files", extensions: ["mbox"] },
      { name: "Outlook PST Files", extensions: ["pst"] },
      { name: "EML Files", extensions: ["eml"] },
      { name: "All Files", extensions: ["*"] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true };
  }

  return openMailboxFile(result.filePaths[0], event.sender);
});

ipcMain.handle("open-mailbox-path", async (event, payload) => {
  const filePath = typeof payload?.filePath === "string" ? payload.filePath : "";
  if (!filePath) {
    return { canceled: true, error: "No mailbox file path was provided." };
  }

  return openMailboxFile(filePath, event.sender);
});

ipcMain.handle("consume-pending-open-file", async () => {
  const filePath = pendingOpenFilePath;
  pendingOpenFilePath = "";
  return { filePath };
});

async function openMailboxFile(filePath, sender) {
  const normalizedFilePath = normalizeMailboxFilePath(filePath);
  if (!normalizedFilePath) {
    return {
      canceled: true,
      error: "Unsupported file type. Please open an .mbox or .pst file."
    };
  }

  const filePathToOpen = normalizedFilePath;
  const openedAsPst = isPstFilePath(filePathToOpen);
  const openedAsEml = isEmlFilePath(filePathToOpen);
  let sourcePath = filePathToOpen;
  let pstConversion = null;

  if (openedAsEml) {
    return openEmlFile(filePathToOpen);
  }

  const sourceStats = await stat(filePathToOpen);

  emitOpenProgress(sender, {
    phase: "preparing",
    filePath: filePathToOpen,
    totalBytes: 0,
    bytesRead: 0,
    messagesIndexed: 0
  });

  if (openedAsPst) {
    const pstDbPath = `${filePathToOpen}.sqlite`;
    const reusableDatabase = await getReusableDatabaseInfo(pstDbPath, filePathToOpen);
    if (reusableDatabase?.sourceEmbedded) {
      await cleanupPstSidecarArtifacts(filePathToOpen);
      emitOpenProgress(sender, {
        phase: "ready",
        filePath: filePathToOpen,
        dbPath: reusableDatabase.dbPath,
        totalBytes: sourceStats.size,
        bytesRead: sourceStats.size,
        messagesIndexed: reusableDatabase.totalMessages,
        reused: true
      });

      const firstPage = searchMessages(reusableDatabase.dbPath, "", DEFAULT_PAGE_SIZE, 0);
      const dateBounds = getMessageDateBounds(reusableDatabase.dbPath);

      return {
        canceled: false,
        filePath: filePathToOpen,
        sourcePath: filePathToOpen,
        sourceType: "pst",
        pstConversion: null,
        dbPath: reusableDatabase.dbPath,
        total: reusableDatabase.totalMessages,
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
    }
  }

  if (openedAsPst) {
    pstConversion = await ensurePstConvertedToMbox(filePathToOpen, {
      onProgress: (payload) => {
        emitOpenProgress(sender, {
          filePath: filePathToOpen,
          ...payload
        });
      }
    });
    sourcePath = pstConversion.mboxPath;
  }

  const indexing = await ensureMboxDatabase(sourcePath, sender, openedAsPst
    ? {
        dbPath: `${filePathToOpen}.sqlite`,
        sourcePath: filePathToOpen,
        persistSourceChunks: true
      }
    : undefined);

  if (openedAsPst) {
    await cleanupPstSidecarArtifacts(filePathToOpen);
  }

  const firstPage = searchMessages(indexing.dbPath, "", DEFAULT_PAGE_SIZE, 0);
  const dateBounds = getMessageDateBounds(indexing.dbPath);

  return {
    canceled: false,
    filePath: filePathToOpen,
    sourcePath: filePathToOpen,
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
}

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

async function cleanupPstSidecarArtifacts(pstPath) {
  const mboxPath = `${pstPath}.mbox`;
  const metaPath = `${mboxPath}.meta.json`;

  await Promise.all(
    [mboxPath, metaPath].map(async (filePath) => {
      try {
        await rm(filePath);
      } catch (error) {
        if (error && error.code !== "ENOENT") {
          console.error(`Failed to remove PST sidecar artifact: ${filePath}`, error);
        }
      }
    })
  );
}

function normalizeMailboxFilePath(filePath) {
  const value = String(filePath || "").trim();
  if (!value || value.startsWith("-")) {
    return "";
  }

  const extension = path.extname(value).toLowerCase();
  if (extension !== ".mbox" && extension !== ".pst" && extension !== ".eml") {
    return "";
  }

  return path.resolve(value);
}

function isEmlFilePath(filePath) {
  return String(filePath || "").toLowerCase().endsWith(".eml");
}

async function openEmlFile(filePath) {
  const raw = await readFile(filePath, "utf8");
  const parsed = parseMessageChunk(raw, {
    index: 1,
    includeAttachmentData: true,
    includeEmlSource: true,
    includeBodyHtml: true
  });

  if (!parsed) {
    return {
      canceled: true,
      error: "The EML file could not be parsed."
    };
  }

  const message = {
    ...parsed,
    id: parsed.id || "eml-1",
    resultIndex: 1
  };

  return {
    canceled: false,
    filePath,
    sourcePath: filePath,
    sourceType: "eml",
    dbPath: "",
    total: 1,
    messages: [
      {
        id: message.id,
        subject: message.subject || "(No Subject)",
        from: message.from || "",
        to: message.to || "",
        date: message.date || "",
        snippet: message.snippet || "",
        hasAttachments: Array.isArray(message.attachments) && message.attachments.length > 0,
        resultIndex: 1
      }
    ],
    offset: 0,
    limit: 1,
    resultTotal: 1,
    dateRange: null,
    standaloneMessage: message
  };
}

function findMailboxFilePathInArgv(argvValues) {
  const values = Array.isArray(argvValues) ? argvValues : [];
  for (let index = values.length - 1; index >= 0; index -= 1) {
    const candidate = normalizeMailboxFilePath(values[index]);
    if (candidate) {
      return candidate;
    }
  }
  return "";
}

function queueOrDispatchMailboxOpen(filePath) {
  const normalizedPath = normalizeMailboxFilePath(filePath);
  if (!normalizedPath) {
    return false;
  }

  if (mainWindow && !mainWindow.isDestroyed() && !mainWindow.webContents.isLoadingMainFrame()) {
    mainWindow.webContents.send(OPEN_MAILBOX_REQUEST_EVENT, { filePath: normalizedPath });
  } else {
    pendingOpenFilePath = normalizedPath;
  }

  return true;
}

function flushPendingMailboxOpenRequest(window) {
  if (!window || window.isDestroyed() || !pendingOpenFilePath) {
    return;
  }

  window.webContents.send(OPEN_MAILBOX_REQUEST_EVENT, { filePath: pendingOpenFilePath });
  pendingOpenFilePath = "";
}

function focusMainWindow() {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return;
  }

  if (mainWindow.isMinimized()) {
    mainWindow.restore();
  }
  mainWindow.show();
  mainWindow.focus();
}

const gotSingleInstanceLock = app.requestSingleInstanceLock();
if (!gotSingleInstanceLock) {
  app.quit();
} else {
  const initialMailboxFilePath = findMailboxFilePathInArgv(process.argv);
  if (initialMailboxFilePath) {
    pendingOpenFilePath = initialMailboxFilePath;
  }

  app.on("second-instance", (_event, commandLine) => {
    const nextFilePath = findMailboxFilePathInArgv(commandLine);
    if (nextFilePath) {
      queueOrDispatchMailboxOpen(nextFilePath);
    }
    focusMainWindow();
  });

  app.on("open-file", (event, filePath) => {
    event.preventDefault();
    queueOrDispatchMailboxOpen(filePath);
    focusMainWindow();
  });

  app.whenReady().then(() => {
    createWindow();

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
        return;
      }
      focusMainWindow();
    });
  });

  app.on("window-all-closed", () => {
    app.quit();
  });
}
