const path = require("path");
const os = require("os");
const { mkdtemp, open, readFile, rm, writeFile } = require("fs/promises");
const { pathToFileURL } = require("url");
const { app, BrowserWindow, Menu, clipboard, dialog, ipcMain, screen, shell } = require("electron");
const {
  ensureMboxDatabase,
  ensurePstDatabase,
  searchMessages,
  loadMessageById,
  getAttachmentData,
  getMessageEmlBuffer,
  getMessageSourcePreview,
  getMessageDateBounds,
  listBookmarkedMessages,
  setMessageBookmarked
} = require("./src/mboxStore");
const { parseMessageChunk } = require("./src/mboxParser");
const { isPstFilePath } = require("./src/pstConverter");
const { createOpenTiming } = require("./src/openTiming");

const DEFAULT_PAGE_SIZE = 200;
const OPEN_PROGRESS_EVENT = "mbox-index-progress";
const OPEN_MAILBOX_REQUEST_EVENT = "open-mailbox-request";
const APP_MENU_COMMAND_EVENT = "app-menu-command";
const PREVIEW_WINDOW_DEFAULT_BOUNDS = { width: 960, height: 760 };
const PREVIEW_WINDOW_MIN_BOUNDS = { width: 480, height: 360 };
const DEFAULT_TOOLBAR_MENU_STATE = Object.freeze({
  canExportBookmarks: false,
  canRemoteContent: false,
  remoteContentEnabled: false,
  canSearch: false,
  canDateFilter: false,
  canFromFilter: false,
  canToFilter: false,
  canSubjectFilter: false,
  canAttachmentFilter: false,
  attachmentsOnlyEnabled: false,
  canBookmarkFilter: false,
  bookmarkedOnlyEnabled: false
});
let attachmentPreviewWindow = null;
let attachmentPreviewWindowState = null;
let mainWindow = null;
let pendingOpenFilePath = "";
let toolbarMenuState = { ...DEFAULT_TOOLBAR_MENU_STATE };

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
      toolbarMenuState = { ...DEFAULT_TOOLBAR_MENU_STATE };
      rebuildApplicationMenu();
    }
  });

  window.webContents.on("did-finish-load", () => {
    flushPendingMailboxOpenRequest(window);
  });

  window.loadFile(path.join(__dirname, "src/renderer/index.html"));
  mainWindow = window;
  rebuildApplicationMenu();
  return window;
}

function rebuildApplicationMenu() {
  if (!app.isReady()) {
    return;
  }
  Menu.setApplicationMenu(Menu.buildFromTemplate(buildApplicationMenuTemplate()));
}

function buildApplicationMenuTemplate() {
  const template = [];
  const appName = app.name || "Mbox Viewer";

  if (process.platform === "darwin") {
    template.push({
      label: appName,
      submenu: [
        { role: "about" },
        { type: "separator" },
        { role: "services" },
        { type: "separator" },
        { role: "hide" },
        { role: "hideOthers" },
        { role: "unhide" },
        { type: "separator" },
        { role: "quit" }
      ]
    });
  }

  template.push(
    {
      label: "File",
      submenu: [
        {
          label: "Open Email or Mailbox...",
          accelerator: "CmdOrCtrl+O",
          click: () => dispatchAppMenuCommand("open-mailbox")
        },
        {
          label: "Export Bookmarked Messages...",
          enabled: toolbarMenuState.canExportBookmarks,
          click: () => dispatchAppMenuCommand("export-bookmarked-mbox")
        },
        { type: "separator" },
        process.platform === "darwin" ? { role: "close" } : { role: "quit" }
      ]
    },
    {
      label: "Edit",
      submenu: [
        { role: "undo" },
        { role: "redo" },
        { type: "separator" },
        { role: "cut" },
        { role: "copy" },
        { role: "paste" },
        { role: "selectAll" }
      ]
    },
    {
      label: "Actions",
      submenu: [
        {
          label: "Remote Content",
          type: "checkbox",
          enabled: toolbarMenuState.canRemoteContent,
          checked: toolbarMenuState.remoteContentEnabled,
          click: () => dispatchAppMenuCommand("toggle-remote-content")
        }
      ]
    },
    {
      label: "Filter",
      submenu: [
        {
          label: "Search Text...",
          accelerator: "CmdOrCtrl+F",
          enabled: toolbarMenuState.canSearch,
          click: () => dispatchAppMenuCommand("focus-search")
        },
        {
          label: "Date Range...",
          enabled: toolbarMenuState.canDateFilter,
          click: () => dispatchAppMenuCommand("open-date-filter")
        },
        {
          label: "Sender...",
          enabled: toolbarMenuState.canFromFilter,
          click: () => dispatchAppMenuCommand("open-from-filter")
        },
        {
          label: "Recipient...",
          enabled: toolbarMenuState.canToFilter,
          click: () => dispatchAppMenuCommand("open-to-filter")
        },
        {
          label: "Subject...",
          enabled: toolbarMenuState.canSubjectFilter,
          click: () => dispatchAppMenuCommand("open-subject-filter")
        },
        { type: "separator" },
        {
          label: "Only Messages With Attachments",
          type: "checkbox",
          enabled: toolbarMenuState.canAttachmentFilter,
          checked: toolbarMenuState.attachmentsOnlyEnabled,
          click: () => dispatchAppMenuCommand("toggle-attachments-filter")
        },
        {
          label: "Only Bookmarked Messages",
          type: "checkbox",
          enabled: toolbarMenuState.canBookmarkFilter,
          checked: toolbarMenuState.bookmarkedOnlyEnabled,
          click: () => dispatchAppMenuCommand("toggle-bookmarked-filter")
        }
      ]
    },
    {
      label: "Window",
      submenu: [
        { role: "minimize" },
        { role: "zoom" },
        ...(process.platform === "darwin" ? [{ type: "separator" }, { role: "front" }] : [{ role: "close" }])
      ]
    }
  );

  return template;
}

function dispatchAppMenuCommand(command) {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return;
  }
  mainWindow.webContents.send(APP_MENU_COMMAND_EVENT, { command });
}

function normalizeToolbarMenuState(payload) {
  const next = payload && typeof payload === "object" ? payload : {};
  return {
    canExportBookmarks: Boolean(next.canExportBookmarks),
    canRemoteContent: Boolean(next.canRemoteContent),
    remoteContentEnabled: Boolean(next.remoteContentEnabled),
    canSearch: Boolean(next.canSearch),
    canDateFilter: Boolean(next.canDateFilter),
    canFromFilter: Boolean(next.canFromFilter),
    canToFilter: Boolean(next.canToFilter),
    canSubjectFilter: Boolean(next.canSubjectFilter),
    canAttachmentFilter: Boolean(next.canAttachmentFilter),
    attachmentsOnlyEnabled: Boolean(next.attachmentsOnlyEnabled),
    canBookmarkFilter: Boolean(next.canBookmarkFilter),
    bookmarkedOnlyEnabled: Boolean(next.bookmarkedOnlyEnabled)
  };
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

ipcMain.handle("update-toolbar-menu-state", async (_, payload) => {
  toolbarMenuState = normalizeToolbarMenuState(payload);
  rebuildApplicationMenu();
  return { ok: true };
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
  const openTiming = createOpenTiming(filePathToOpen);
  openTiming.mark("open-start", { sourceType: openedAsPst ? "pst" : openedAsEml ? "eml" : "mbox" });

  if (openedAsEml) {
    return openEmlFile(filePathToOpen, openTiming);
  }
  const indexing = openedAsPst
    ? await ensurePstDatabase(filePathToOpen, sender, { dbPath: `${filePathToOpen}.sqlite` })
    : await ensureMboxDatabase(filePathToOpen, sender);
  openTiming.mark("index-ready", {
    reused: Boolean(indexing?.reused),
    totalMessages: Number(indexing?.totalMessages) || 0
  });

  const firstPage = searchMessages(indexing.dbPath, "", DEFAULT_PAGE_SIZE, 0);
  const dateBounds = getMessageDateBounds(indexing.dbPath);
  openTiming.mark("first-page-ready", {
    pageCount: Array.isArray(firstPage.messages) ? firstPage.messages.length : 0
  });

  return {
    canceled: false,
    filePath: filePathToOpen,
    sourcePath: filePathToOpen,
    sourceType: openedAsPst ? "pst" : "mbox",
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
      : null,
    openTiming: openTiming.snapshot({
      reused: Boolean(indexing?.reused),
      totalMessages: Number(indexing?.totalMessages) || 0
    })
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
  const bookmarkedOnly = Boolean(payload?.bookmarkedOnly);

  if (!dbPath) {
    return { total: 0, offset: 0, limit: 0, messages: [] };
  }

  return searchMessages(dbPath, query, limit, offset, {
    dateFrom,
    dateTo,
    senderQuery,
    recipientQuery,
    subjectQuery,
    attachmentsOnly,
    bookmarkedOnly
  });
});

ipcMain.handle("set-message-bookmarked", async (_, payload) => {
  const dbPath = typeof payload?.dbPath === "string" ? payload.dbPath : "";
  const id = payload?.id;
  const isBookmarked = Boolean(payload?.isBookmarked);

  if (!dbPath || id === undefined || id === null) {
    return null;
  }

  return setMessageBookmarked(dbPath, id, isBookmarked);
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
  const dbPath = typeof payload?.dbPath === "string" ? payload.dbPath : "";
  const messageId = payload?.messageId;
  const attachmentId = typeof payload?.attachmentId === "string" ? payload.attachmentId : "";
  const fileName = typeof payload?.fileName === "string" ? payload.fileName : "attachment.bin";
  const base64 = typeof payload?.base64 === "string" ? payload.base64 : "";
  const attachmentData =
    base64
      ? {
          fileName,
          contentType: typeof payload?.contentType === "string" ? payload.contentType : "application/octet-stream",
          data: Buffer.from(base64, "base64")
        }
      : dbPath && messageId != null && attachmentId
        ? await getAttachmentData(dbPath, messageId, attachmentId)
        : null;

  if (!attachmentData?.data) {
    return { canceled: true, error: "No attachment data" };
  }

  const result = await dialog.showSaveDialog({
    title: "Save attachment",
    defaultPath: attachmentData.fileName || fileName
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  await writeFile(result.filePath, attachmentData.data);
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle("save-message-eml", async (_, payload) => {
  const rawFileName = typeof payload?.fileName === "string" ? payload.fileName : "message.eml";
  const fileName = normalizeEmlFileName(rawFileName);
  const dbPath = typeof payload?.dbPath === "string" ? payload.dbPath : "";
  const messageId = payload?.messageId;
  const sourcePath = typeof payload?.sourcePath === "string" ? payload.sourcePath : "";
  const emlBuffer =
    dbPath && messageId != null
      ? await getMessageEmlBuffer(dbPath, messageId)
      : sourcePath
        ? await readFile(sourcePath)
        : null;

  if (!emlBuffer || emlBuffer.length === 0) {
    return { canceled: true, error: "No message source available" };
  }

  const result = await dialog.showSaveDialog({
    title: "Save message as EML",
    defaultPath: fileName
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  await writeFile(result.filePath, emlBuffer);
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle("export-bookmarked-mbox", async (_, payload) => {
  const dbPath = typeof payload?.dbPath === "string" ? payload.dbPath : "";
  const sourcePath = typeof payload?.sourcePath === "string" ? payload.sourcePath : "";

  if (!dbPath) {
    return { canceled: true, error: "No mailbox database is available." };
  }

  const bookmarkedMessages = listBookmarkedMessages(dbPath);
  if (!Array.isArray(bookmarkedMessages) || bookmarkedMessages.length === 0) {
    return { canceled: true, error: "No bookmarked messages available." };
  }

  const defaultBaseName = buildBookmarkedMboxFileName(sourcePath || dbPath);
  const result = await dialog.showSaveDialog({
    title: "Export bookmarked messages",
    defaultPath: defaultBaseName,
    filters: [{ name: "Mbox Files", extensions: ["mbox"] }]
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  const handle = await open(result.filePath, "w");
  let exportedCount = 0;
  try {
    for (const message of bookmarkedMessages) {
      const emlBuffer = await getMessageEmlBuffer(dbPath, message.id);
      if (!emlBuffer || emlBuffer.length === 0) {
        continue;
      }

      const entryBuffer = buildMboxEntryBuffer(message, emlBuffer);
      await handle.write(entryBuffer);
      exportedCount += 1;
    }
  } finally {
    await handle.close();
  }

  return {
    canceled: false,
    filePath: result.filePath,
    exportedCount
  };
});

ipcMain.handle("get-message-source-preview", async (_, payload) => {
  const dbPath = typeof payload?.dbPath === "string" ? payload.dbPath : "";
  const messageId = payload?.messageId;
  const sourcePath = typeof payload?.sourcePath === "string" ? payload.sourcePath : "";

  if (dbPath && messageId != null) {
    return { text: await getMessageSourcePreview(dbPath, messageId) };
  }

  if (sourcePath) {
    const raw = await readFile(sourcePath);
    return { text: raw.toString("utf8") };
  }

  return { text: "" };
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
  const dbPath = typeof payload?.dbPath === "string" ? payload.dbPath : "";
  const messageId = payload?.messageId;
  const attachmentId = typeof payload?.attachmentId === "string" ? payload.attachmentId : "";
  const base64 = typeof payload?.base64 === "string" ? payload.base64 : "";
  const contentType = typeof payload?.contentType === "string" ? payload.contentType : "";
  const fileName = typeof payload?.fileName === "string" ? payload.fileName : "attachment";
  const attachmentData =
    base64
      ? {
          fileName,
          contentType,
          data: Buffer.from(base64, "base64")
        }
      : dbPath && messageId != null && attachmentId
        ? await getAttachmentData(dbPath, messageId, attachmentId)
        : null;
  let tempDir = "";

  if (!attachmentData?.data || !isPreviewableContentType(attachmentData.contentType)) {
    return { opened: false };
  }

  try {
    tempDir = await mkdtemp(path.join(os.tmpdir(), "mbox-viewer-preview-"));
    const tempFilePath = path.join(
      tempDir,
      buildPreviewFileName(attachmentData.fileName || fileName, attachmentData.contentType)
    );
    await writeFile(tempFilePath, attachmentData.data);

    const parentWindow = BrowserWindow.fromWebContents(event.sender) || null;
    const previewWindow = await createAttachmentPreviewWindow(parentWindow);
    const previewUrl = buildAttachmentPreviewPageUrl(
      tempFilePath,
      attachmentData.fileName || fileName,
      attachmentData.contentType
    );

    await previewWindow.loadURL(previewUrl);
    previewWindow.setTitle(attachmentData.fileName || fileName || "Attachment Preview");
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

function normalizeMboxFileName(input) {
  const stripped = String(input || "bookmarked-messages")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  const base = stripped || "bookmarked-messages";
  return base.toLowerCase().endsWith(".mbox") ? base : `${base}.mbox`;
}

function buildBookmarkedMboxFileName(sourcePath) {
  const sourceName = path.basename(String(sourcePath || ""), path.extname(String(sourcePath || "")));
  return normalizeMboxFileName(sourceName ? `${sourceName}-bookmarked` : "bookmarked-messages");
}

function buildMboxEntryBuffer(message, emlBuffer) {
  const envelopeFrom = getMboxEnvelopeAddress(message?.from);
  const envelopeDate = formatMboxEnvelopeDate(message?.date);
  const headerBuffer = Buffer.from(`From ${envelopeFrom} ${envelopeDate}\n`, "utf8");
  const escapedMessageBuffer = escapeMessageForMbox(emlBuffer);
  const needsTrailingNewline =
    escapedMessageBuffer.length === 0 ||
    escapedMessageBuffer[escapedMessageBuffer.length - 1] !== 0x0a;

  return Buffer.concat(
    [
      headerBuffer,
      escapedMessageBuffer,
      needsTrailingNewline ? Buffer.from("\n", "utf8") : Buffer.alloc(0),
      Buffer.from("\n", "utf8")
    ].filter((part) => part.length > 0)
  );
}

function getMboxEnvelopeAddress(fromValue) {
  const raw = String(fromValue || "").trim();
  const address =
    raw.match(/<([^<>\s]+@[^<>\s]+)>/)?.[1] ||
    raw.match(/\b([^\s<>@]+@[^\s<>@]+)\b/)?.[1] ||
    "unknown@mboxviewer.local";
  return address.replace(/\s+/g, "");
}

function formatMboxEnvelopeDate(dateValue) {
  const timestamp = Date.parse(String(dateValue || ""));
  const date = Number.isFinite(timestamp) ? new Date(timestamp) : new Date();
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const dayName = days[date.getDay()];
  const monthName = months[date.getMonth()];
  const dayOfMonth = String(date.getDate()).padStart(2, " ");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");
  const year = date.getFullYear();
  return `${dayName} ${monthName} ${dayOfMonth} ${hours}:${minutes}:${seconds} ${year}`;
}

function escapeMessageForMbox(buffer) {
  const source = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer || "");
  const parts = [];
  let cursor = 0;

  while (cursor < source.length) {
    const newlineIndex = source.indexOf(0x0a, cursor);
    const lineEnd = newlineIndex === -1 ? source.length : newlineIndex + 1;
    const line = source.subarray(cursor, lineEnd);
    if (
      line.length >= 5 &&
      line[0] === 0x46 &&
      line[1] === 0x72 &&
      line[2] === 0x6f &&
      line[3] === 0x6d &&
      line[4] === 0x20
    ) {
      parts.push(Buffer.from(">", "utf8"));
    }
    parts.push(line);
    cursor = lineEnd;
  }

  return Buffer.concat(parts);
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

async function openEmlFile(filePath, openTiming = null) {
  const raw = await readFile(filePath);
  const parsed = parseMessageChunk(raw, {
    index: 1,
    includeAttachmentData: "all",
    includeEmlSource: false,
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
    id: 1,
    sourcePath: filePath,
    resultIndex: 1
  };
  openTiming?.mark("first-page-ready", { pageCount: 1 });

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
    standaloneMessage: message,
    openTiming: openTiming?.snapshot({
      reused: false,
      totalMessages: 1
    }) || null
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
