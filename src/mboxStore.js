const { stat, open, rm } = require("fs/promises");
const { createReadStream } = require("fs");
const { setImmediate: waitForImmediate } = require("timers/promises");
const { parseMessageChunk } = require("./mboxParser");

const SCHEMA_VERSION = "2";
const PARSER_VERSION = "1";
const PROGRESS_EVENT = "mbox-index-progress";
const DB_CACHE = new Map();
const BATCH_SIZE = 200;
const MESSAGE_LIST_ORDER_SQL = "date_ts IS NULL, date_ts DESC, id DESC";
const MESSAGE_LIST_ORDER_SQL_ALIASED = "m.date_ts IS NULL, m.date_ts DESC, m.id DESC";
const DatabaseClass = resolveDatabaseClass();

function resolveDatabaseClass() {
  try {
    const { DatabaseSync } = require("node:sqlite");
    if (typeof DatabaseSync === "function") {
      return DatabaseSync;
    }
  } catch {
    // Fallback to external SQLite package for Electron runtimes without node:sqlite.
  }

  try {
    const BetterSqlite3 = require("better-sqlite3");
    if (typeof BetterSqlite3 === "function") {
      return BetterSqlite3;
    }
  } catch (error) {
    throw new Error(
      `No SQLite runtime available. Install better-sqlite3 and rebuild native deps for Electron. ${error.message}`
    );
  }

  throw new Error("Unsupported SQLite runtime.");
}

async function ensureMboxDatabase(filePath, sender) {
  const sourceStats = await stat(filePath);
  const dbPath = `${filePath}.sqlite`;
  const sourceMtimeMs = Math.trunc(sourceStats.mtimeMs);

  emitProgress(sender, {
    phase: "preparing",
    filePath,
    dbPath,
    totalBytes: sourceStats.size,
    bytesRead: 0,
    messagesIndexed: 0
  });

  if (await isReusableDatabase(dbPath, filePath, sourceStats.size, sourceMtimeMs)) {
    const totalMessages = countMessages(dbPath);
    emitProgress(sender, {
      phase: "ready",
      filePath,
      dbPath,
      totalBytes: sourceStats.size,
      bytesRead: sourceStats.size,
      messagesIndexed: totalMessages,
      reused: true
    });
    return { dbPath, totalMessages, reused: true };
  }

  closeDatabase(dbPath);
  await removeDbFiles(dbPath);
  const db = createWritableDatabase(dbPath);

  const insertMessage = db.prepare(`
    INSERT INTO messages (
      id,
      subject,
      sender,
      recipient,
      date_raw,
      date_ts,
      snippet,
      body_text,
      attachment_names,
      byte_start,
      byte_end
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const insertAttachment = db.prepare(`
    INSERT INTO attachments (
      message_id,
      file_name,
      content_type,
      size,
      is_inline,
      content_id
    ) VALUES (?, ?, ?, ?, ?, ?)
  `);
  const insertFts = db.prepare(`
    INSERT INTO message_fts (
      rowid,
      subject,
      sender,
      recipient,
      snippet,
      body_text,
      attachment_names
    ) VALUES (?, ?, ?, ?, ?, ?, ?)
  `);

  let messagesIndexed = 0;
  let bytesRead = 0;
  let lastProgressAt = 0;
  let transactionOpen = false;
  let writableDbClosed = false;

  const closeWritableDb = () => {
    if (writableDbClosed) {
      return;
    }
    writableDbClosed = true;
    try {
      db.close();
    } catch {
      // Ignore close failures during teardown.
    }
  };

  const commitBatch = () => {
    if (!transactionOpen) {
      return;
    }
    db.exec("COMMIT");
    transactionOpen = false;
  };

  try {
    db.exec("BEGIN");
    transactionOpen = true;

    await streamMboxMessages(
      filePath,
      async ({ rawChunk, byteStart, byteEnd }) => {
        messagesIndexed += 1;
        const parsed = parseMessageChunk(rawChunk, {
          index: messagesIndexed,
          includeAttachmentData: false,
          includeEmlSource: false,
          includeBodyHtml: false
        });

        if (!parsed) {
          return;
        }

        const subject = parsed.subject || "";
        const senderValue = parsed.from || "";
        const recipient = parsed.to || "";
        const dateRaw = parsed.date || "";
        const snippet = parsed.snippet || "";
        const dateTs = parseMessageDateToTimestamp(dateRaw);
        const bodyText = parsed.bodyText || "";
        const attachmentNames = (parsed.attachments || [])
          .map((attachment) => attachment.fileName || "")
          .filter(Boolean)
          .join(" ");

        insertMessage.run(
          messagesIndexed,
          subject,
          senderValue,
          recipient,
          dateRaw,
          dateTs,
          snippet,
          bodyText,
          attachmentNames,
          byteStart,
          byteEnd
        );

        insertFts.run(
          messagesIndexed,
          subject,
          senderValue,
          recipient,
          snippet,
          bodyText,
          attachmentNames
        );

        for (const attachment of parsed.attachments || []) {
          insertAttachment.run(
            messagesIndexed,
            attachment.fileName || "",
            attachment.contentType || "",
            attachment.size,
            attachment.isInline ? 1 : 0,
            attachment.contentId || ""
          );
        }

        if (messagesIndexed % BATCH_SIZE === 0) {
          commitBatch();
          db.exec("BEGIN");
          transactionOpen = true;
          await waitForImmediate();
        }
      },
      async (progress) => {
        bytesRead = progress.bytesRead;
        const now = Date.now();
        if (progress.done || now - lastProgressAt >= 200) {
          lastProgressAt = now;
          emitProgress(sender, {
            phase: "indexing",
            filePath,
            dbPath,
            totalBytes: sourceStats.size,
            bytesRead,
            messagesIndexed
          });
          await waitForImmediate();
        }
      }
    );

    commitBatch();
    db.exec("ANALYZE");
    db.exec("INSERT INTO message_fts(message_fts) VALUES ('optimize')");
    writeMeta(db, filePath, sourceStats.size, sourceMtimeMs, messagesIndexed);
  } catch (error) {
    if (transactionOpen) {
      try {
        db.exec("ROLLBACK");
      } catch {
        // Ignore rollback failures after parser/db errors.
      }
    }
    closeWritableDb();
    throw error;
  }

  closeWritableDb();
  emitProgress(sender, {
    phase: "ready",
    filePath,
    dbPath,
    totalBytes: sourceStats.size,
    bytesRead: sourceStats.size,
    messagesIndexed,
    reused: false
  });

  return { dbPath, totalMessages: messagesIndexed, reused: false };
}

function searchMessages(dbPath, queryInput, limitInput, offsetInput, filtersInput = null) {
  const entry = getDatabaseEntry(dbPath);
  const query = String(queryInput || "").trim();
  const limit = clampNumber(limitInput, 1, 500, 200);
  const offset = clampNumber(offsetInput, 0, Number.MAX_SAFE_INTEGER, 0);
  const dateRange = normalizeDateRange(filtersInput);
  const fieldFilters = normalizeFieldFilters(filtersInput);
  const hasDateFilter = dateRange !== null;
  const hasFieldFilters = fieldFilters !== null;

  let rows = [];
  let total = 0;

  if (!query && !hasDateFilter && !hasFieldFilters) {
    total = entry.countAll.get().count;
    rows = entry.listAll.all(limit, offset);
  } else if (!query && !hasFieldFilters) {
    total = entry.countAllByDate.get(dateRange.from, dateRange.to).count;
    rows = entry.listAllByDate.all(dateRange.from, dateRange.to, limit, offset);
  } else {
    const searchSpec = buildMessageSearchSpec({
      query,
      dateRange,
      fieldFilters
    });
    if (!searchSpec) {
      return { total: 0, offset, limit, messages: [] };
    }
    total = entry.db.prepare(searchSpec.countSql).get(...searchSpec.countParams).count;
    rows = entry.db.prepare(searchSpec.listSql).all(...searchSpec.listParams, limit, offset);
  }

  const messages = rows.map((row, index) => ({
    id: row.id,
    subject: row.subject || "(No Subject)",
    from: row.sender || "",
    to: row.recipient || "",
    date: row.date_raw || "",
    snippet: row.snippet || "",
    hasAttachments: Boolean(row.has_attachments),
    resultIndex: offset + index + 1
  }));

  return { total, offset, limit, messages };
}

function getMessageDateBounds(dbPath) {
  const entry = getDatabaseEntry(dbPath);
  const row = entry.getDateBounds.get();
  if (!row || row.min_date_ts === null || row.max_date_ts === null) {
    return null;
  }

  return {
    minDateTs: Number(row.min_date_ts),
    maxDateTs: Number(row.max_date_ts),
    datedCount: Number(row.dated_count) || 0
  };
}

async function loadMessageById(dbPath, messageId) {
  const entry = getDatabaseEntry(dbPath);
  const id = Number.parseInt(String(messageId), 10);
  if (!Number.isInteger(id) || id <= 0) {
    return null;
  }

  const row = entry.getMessage.get(id);
  if (!row) {
    return null;
  }

  const sourcePath = entry.getMeta.get("source_path")?.value;
  if (!sourcePath) {
    throw new Error("Indexed database is missing source file metadata.");
  }

  const rawChunk = await readUtf8Range(sourcePath, row.byte_start, row.byte_end);
  const parsed = parseMessageChunk(rawChunk, {
    index: row.id,
    includeAttachmentData: true,
    includeEmlSource: true,
    includeBodyHtml: true
  });

  if (!parsed) {
    return null;
  }

  return {
    ...parsed,
    id: row.id,
    subject: parsed.subject || row.subject || "(No Subject)",
    from: parsed.from || row.sender || "",
    to: parsed.to || row.recipient || "",
    date: parsed.date || row.date_raw || "",
    snippet: parsed.snippet || row.snippet || "",
    resultIndex: null
  };
}

function countMessages(dbPath) {
  const entry = getDatabaseEntry(dbPath);
  return entry.countAll.get().count;
}

async function streamMboxMessages(filePath, onMessage, onProgress) {
  const fileStats = await stat(filePath);
  const totalBytes = fileStats.size;

  const input = createReadStream(filePath);
  let carry = Buffer.alloc(0);
  let carryOffset = 0;
  let fileOffset = 0;
  let currentStartOffset = null;
  let currentLines = [];
  let currentLength = 0;

  const maybeProgress = async (bytesRead, done = false) => {
    if (typeof onProgress !== "function") {
      return;
    }
    await onProgress({
      bytesRead: Math.min(bytesRead, totalBytes),
      totalBytes,
      done
    });
  };

  const finalizeCurrentMessage = async (endOffset) => {
    if (currentStartOffset === null) {
      return;
    }
    const rawChunk = Buffer.concat(currentLines, currentLength).toString("utf8");
    await onMessage({
      rawChunk,
      byteStart: currentStartOffset,
      byteEnd: endOffset
    });
  };

  const startMessage = (lineBuffer, lineOffset) => {
    currentStartOffset = lineOffset;
    currentLines = [lineBuffer];
    currentLength = lineBuffer.length;
  };

  const appendToMessage = (lineBuffer) => {
    currentLines.push(lineBuffer);
    currentLength += lineBuffer.length;
  };

  const processLine = async (lineBuffer, lineOffset) => {
    const isBoundary = isFromBoundaryLine(lineBuffer);
    if (isBoundary && currentStartOffset !== null) {
      await finalizeCurrentMessage(lineOffset);
      startMessage(lineBuffer, lineOffset);
    } else if (isBoundary) {
      startMessage(lineBuffer, lineOffset);
    } else if (currentStartOffset !== null) {
      appendToMessage(lineBuffer);
    }

    await maybeProgress(lineOffset + lineBuffer.length, false);
  };

  for await (const chunk of input) {
    const combined = carry.length > 0 ? Buffer.concat([carry, chunk]) : chunk;
    const combinedOffset = carry.length > 0 ? carryOffset : fileOffset;
    let cursor = 0;
    let newlineIndex = combined.indexOf(0x0a, cursor);

    while (newlineIndex !== -1) {
      const lineBuffer = combined.subarray(cursor, newlineIndex + 1);
      const lineOffset = combinedOffset + cursor;
      await processLine(lineBuffer, lineOffset);
      cursor = newlineIndex + 1;
      newlineIndex = combined.indexOf(0x0a, cursor);
    }

    carry = combined.subarray(cursor);
    carryOffset = combinedOffset + cursor;
    fileOffset += chunk.length;
  }

  if (carry.length > 0) {
    await processLine(carry, carryOffset);
  }

  if (currentStartOffset !== null) {
    await finalizeCurrentMessage(totalBytes);
  }

  await maybeProgress(totalBytes, true);
}

function isFromBoundaryLine(lineBuffer) {
  let end = lineBuffer.length;
  if (end > 0 && lineBuffer[end - 1] === 0x0a) {
    end -= 1;
  }
  if (end > 0 && lineBuffer[end - 1] === 0x0d) {
    end -= 1;
  }
  if (end < 5) {
    return false;
  }
  return (
    lineBuffer[0] === 0x46 &&
    lineBuffer[1] === 0x72 &&
    lineBuffer[2] === 0x6f &&
    lineBuffer[3] === 0x6d &&
    lineBuffer[4] === 0x20
  );
}

async function isReusableDatabase(dbPath, sourcePath, sourceSize, sourceMtimeMs) {
  const fileExists = await fileExistsAtPath(dbPath);
  if (!fileExists) {
    return false;
  }

  try {
    const entry = getDatabaseEntry(dbPath);
    const readMeta = (key) => entry.getMeta.get(key)?.value || "";
    const schemaVersion = readMeta("schema_version");
    const parserVersion = readMeta("parser_version");
    const storedSourcePath = readMeta("source_path");
    const storedSourceSize = Number.parseInt(readMeta("source_size"), 10);
    const storedSourceMtime = Number.parseInt(readMeta("source_mtime_ms"), 10);
    const valid =
      schemaVersion === SCHEMA_VERSION &&
      parserVersion === PARSER_VERSION &&
      storedSourcePath === sourcePath &&
      storedSourceSize === sourceSize &&
      storedSourceMtime === sourceMtimeMs;

    if (!valid) {
      closeDatabase(dbPath);
    }

    return valid;
  } catch {
    closeDatabase(dbPath);
    return false;
  }
}

async function removeDbFiles(dbPath) {
  const files = [dbPath, `${dbPath}-wal`, `${dbPath}-shm`];
  for (const filePath of files) {
    try {
      await rm(filePath);
    } catch (error) {
      if (error && error.code !== "ENOENT") {
        throw error;
      }
    }
  }
}

function createWritableDatabase(dbPath) {
  const db = new DatabaseClass(dbPath);
  db.exec("PRAGMA journal_mode=WAL");
  db.exec("PRAGMA synchronous=NORMAL");
  db.exec("PRAGMA temp_store=MEMORY");
  db.exec("PRAGMA foreign_keys=ON");
  db.exec(`
    CREATE TABLE IF NOT EXISTS meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS messages (
      id INTEGER PRIMARY KEY,
      subject TEXT NOT NULL DEFAULT '',
      sender TEXT NOT NULL DEFAULT '',
      recipient TEXT NOT NULL DEFAULT '',
      date_raw TEXT NOT NULL DEFAULT '',
      date_ts INTEGER,
      snippet TEXT NOT NULL DEFAULT '',
      body_text TEXT NOT NULL DEFAULT '',
      attachment_names TEXT NOT NULL DEFAULT '',
      byte_start INTEGER NOT NULL,
      byte_end INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS attachments (
      id INTEGER PRIMARY KEY,
      message_id INTEGER NOT NULL REFERENCES messages(id) ON DELETE CASCADE,
      file_name TEXT NOT NULL DEFAULT '',
      content_type TEXT NOT NULL DEFAULT '',
      size INTEGER,
      is_inline INTEGER NOT NULL DEFAULT 0,
      content_id TEXT NOT NULL DEFAULT ''
    );

    CREATE VIRTUAL TABLE IF NOT EXISTS message_fts USING fts5(
      subject,
      sender,
      recipient,
      snippet,
      body_text,
      attachment_names,
      tokenize='unicode61 remove_diacritics 2'
    );

    CREATE INDEX IF NOT EXISTS idx_messages_date_ts ON messages(date_ts);
  `);
  db.exec("DELETE FROM meta");
  db.exec("DELETE FROM attachments");
  db.exec("DELETE FROM messages");
  db.exec("DELETE FROM message_fts");
  return db;
}

function writeMeta(db, sourcePath, sourceSize, sourceMtimeMs, totalMessages) {
  const upsert = db.prepare(`
    INSERT INTO meta (key, value)
    VALUES (?, ?)
    ON CONFLICT(key) DO UPDATE SET value = excluded.value
  `);
  upsert.run("schema_version", SCHEMA_VERSION);
  upsert.run("parser_version", PARSER_VERSION);
  upsert.run("source_path", sourcePath);
  upsert.run("source_size", String(sourceSize));
  upsert.run("source_mtime_ms", String(sourceMtimeMs));
  upsert.run("total_messages", String(totalMessages));
  upsert.run("indexed_at", new Date().toISOString());
}

function getDatabaseEntry(dbPath) {
  const existing = DB_CACHE.get(dbPath);
  if (existing) {
    return existing;
  }

  const db = new DatabaseClass(dbPath);
  db.exec("PRAGMA foreign_keys=ON");
  const entry = {
    db,
    countAll: db.prepare("SELECT COUNT(*) AS count FROM messages"),
    countAllByDate: db.prepare(`
      SELECT COUNT(*) AS count
      FROM messages
      WHERE date_ts IS NOT NULL
        AND date_ts BETWEEN ? AND ?
    `),
    listAll: db.prepare(`
      SELECT id, subject, sender, recipient, date_raw, snippet, CASE WHEN attachment_names != '' THEN 1 ELSE 0 END AS has_attachments
      FROM messages
      ORDER BY ${MESSAGE_LIST_ORDER_SQL}
      LIMIT ? OFFSET ?
    `),
    listAllByDate: db.prepare(`
      SELECT id, subject, sender, recipient, date_raw, snippet, CASE WHEN attachment_names != '' THEN 1 ELSE 0 END AS has_attachments
      FROM messages
      WHERE date_ts IS NOT NULL
        AND date_ts BETWEEN ? AND ?
      ORDER BY ${MESSAGE_LIST_ORDER_SQL}
      LIMIT ? OFFSET ?
    `),
    countSearch: db.prepare(`
      SELECT COUNT(*) AS count
      FROM message_fts
      WHERE message_fts MATCH ?
    `),
    countSearchByDate: db.prepare(`
      SELECT COUNT(*) AS count
      FROM message_fts
      JOIN messages m ON m.id = message_fts.rowid
      WHERE message_fts MATCH ?
        AND m.date_ts IS NOT NULL
        AND m.date_ts BETWEEN ? AND ?
    `),
    listSearch: db.prepare(`
      SELECT
        m.id,
        m.subject,
        m.sender,
        m.recipient,
        m.date_raw,
        m.snippet,
        CASE WHEN m.attachment_names != '' THEN 1 ELSE 0 END AS has_attachments
      FROM message_fts
      JOIN messages m ON m.id = message_fts.rowid
      WHERE message_fts MATCH ?
      ORDER BY ${MESSAGE_LIST_ORDER_SQL_ALIASED}
      LIMIT ? OFFSET ?
    `),
    listSearchByDate: db.prepare(`
      SELECT
        m.id,
        m.subject,
        m.sender,
        m.recipient,
        m.date_raw,
        m.snippet,
        CASE WHEN m.attachment_names != '' THEN 1 ELSE 0 END AS has_attachments
      FROM message_fts
      JOIN messages m ON m.id = message_fts.rowid
      WHERE message_fts MATCH ?
        AND m.date_ts IS NOT NULL
        AND m.date_ts BETWEEN ? AND ?
      ORDER BY ${MESSAGE_LIST_ORDER_SQL_ALIASED}
      LIMIT ? OFFSET ?
    `),
    getMessage: db.prepare(`
      SELECT id, subject, sender, recipient, date_raw, snippet, byte_start, byte_end
      FROM messages
      WHERE id = ?
    `),
    getMeta: db.prepare("SELECT value FROM meta WHERE key = ?"),
    getDateBounds: db.prepare(`
      SELECT
        MIN(date_ts) AS min_date_ts,
        MAX(date_ts) AS max_date_ts,
        COUNT(date_ts) AS dated_count
      FROM messages
      WHERE date_ts IS NOT NULL
    `)
  };
  DB_CACHE.set(dbPath, entry);
  return entry;
}

function closeDatabase(dbPath) {
  const entry = DB_CACHE.get(dbPath);
  if (!entry) {
    return;
  }

  try {
    entry.db.close();
  } catch {
    // Ignore already-closed handles.
  }
  DB_CACHE.delete(dbPath);
}

async function readUtf8Range(filePath, byteStart, byteEnd) {
  const start = Number.parseInt(String(byteStart), 10);
  const end = Number.parseInt(String(byteEnd), 10);
  if (!Number.isInteger(start) || !Number.isInteger(end) || end < start) {
    throw new Error("Invalid message byte offsets in index.");
  }
  const length = end - start;
  if (length === 0) {
    return "";
  }

  const buffer = Buffer.alloc(length);
  const handle = await open(filePath, "r");
  try {
    const { bytesRead } = await handle.read(buffer, 0, length, start);
    return buffer.subarray(0, bytesRead).toString("utf8");
  } finally {
    await handle.close();
  }
}

function buildFtsQuery(input) {
  const terms = String(input || "")
    .trim()
    .split(/\s+/)
    .map((term) => term.trim())
    .filter(Boolean)
    .slice(0, 16);

  if (terms.length === 0) {
    return "";
  }

  return terms
    .map((term) => {
      const escaped = term.replace(/"/g, '""');
      return `"${escaped}"*`;
    })
    .join(" AND ");
}

function buildMessageSearchSpec({ query, dateRange, fieldFilters }) {
  const ftsQuery = query ? buildFtsQuery(query) : "";
  if (query && !ftsQuery) {
    return null;
  }

  const where = [];
  const params = [];
  const hasFtsQuery = Boolean(ftsQuery);
  const senderQuery = fieldFilters?.senderQuery || "";
  const recipientQuery = fieldFilters?.recipientQuery || "";
  const subjectQuery = fieldFilters?.subjectQuery || "";
  const attachmentsOnly = Boolean(fieldFilters?.attachmentsOnly);

  const fromSql = hasFtsQuery
    ? "FROM message_fts JOIN messages m ON m.id = message_fts.rowid"
    : "FROM messages m";

  if (hasFtsQuery) {
    where.push("message_fts MATCH ?");
    params.push(ftsQuery);
  }

  if (dateRange) {
    where.push("m.date_ts IS NOT NULL");
    where.push("m.date_ts BETWEEN ? AND ?");
    params.push(dateRange.from, dateRange.to);
  }

  if (senderQuery) {
    where.push("m.sender LIKE ? ESCAPE '\\' COLLATE NOCASE");
    params.push(buildLikeContainsPattern(senderQuery));
  }

  if (recipientQuery) {
    where.push("m.recipient LIKE ? ESCAPE '\\' COLLATE NOCASE");
    params.push(buildLikeContainsPattern(recipientQuery));
  }

  if (subjectQuery) {
    where.push("m.subject LIKE ? ESCAPE '\\' COLLATE NOCASE");
    params.push(buildLikeContainsPattern(subjectQuery));
  }

  if (attachmentsOnly) {
    where.push("m.attachment_names != ''");
  }

  const whereSql = where.length > 0 ? ` WHERE ${where.join(" AND ")}` : "";
  const orderSql = ` ORDER BY ${MESSAGE_LIST_ORDER_SQL_ALIASED}`;

  return {
    countSql: `SELECT COUNT(*) AS count ${fromSql}${whereSql}`,
    countParams: params,
    listSql: `
      SELECT
        m.id,
        m.subject,
        m.sender,
        m.recipient,
        m.date_raw,
        m.snippet,
        CASE WHEN m.attachment_names != '' THEN 1 ELSE 0 END AS has_attachments
      ${fromSql}${whereSql}${orderSql}
      LIMIT ? OFFSET ?
    `,
    listParams: params
  };
}

function normalizeDateRange(filtersInput) {
  const from = normalizeTimestamp(filtersInput?.dateFrom);
  const to = normalizeTimestamp(filtersInput?.dateTo);
  if (from === null && to === null) {
    return null;
  }

  const minBoundary = Number.MIN_SAFE_INTEGER;
  const maxBoundary = Number.MAX_SAFE_INTEGER;
  let normalizedFrom = from === null ? minBoundary : from;
  let normalizedTo = to === null ? maxBoundary : to;

  if (normalizedFrom > normalizedTo) {
    const temp = normalizedFrom;
    normalizedFrom = normalizedTo;
    normalizedTo = temp;
  }

  return { from: normalizedFrom, to: normalizedTo };
}

function normalizeFieldFilters(filtersInput) {
  const senderQuery = normalizeSearchFilterValue(filtersInput?.senderQuery);
  const recipientQuery = normalizeSearchFilterValue(filtersInput?.recipientQuery);
  const subjectQuery = normalizeSearchFilterValue(filtersInput?.subjectQuery);
  const attachmentsOnly = Boolean(filtersInput?.attachmentsOnly);

  if (!senderQuery && !recipientQuery && !subjectQuery && !attachmentsOnly) {
    return null;
  }

  return {
    senderQuery,
    recipientQuery,
    subjectQuery,
    attachmentsOnly
  };
}

function normalizeSearchFilterValue(value) {
  return String(value || "").trim();
}

function buildLikeContainsPattern(value) {
  const escaped = String(value || "").replace(/[\\%_]/g, "\\$&");
  return `%${escaped}%`;
}

function normalizeTimestamp(value) {
  if (value === undefined || value === null || value === "") {
    return null;
  }

  const parsed = Number.parseInt(String(value), 10);
  if (!Number.isFinite(parsed)) {
    return null;
  }

  return parsed;
}

function parseMessageDateToTimestamp(dateValue) {
  const raw = String(dateValue || "").trim();
  if (!raw) {
    return null;
  }

  const candidates = new Set([
    raw,
    raw.replace(/\s*\([^)]*\)\s*$/, ""),
    raw.replace(/\s*\([^)]*\)\s*/g, " ").replace(/\s+/g, " ").trim()
  ]);

  for (const candidate of candidates) {
    if (!candidate) {
      continue;
    }

    const timestamp = Date.parse(candidate);
    if (Number.isFinite(timestamp)) {
      return timestamp;
    }
  }

  return null;
}

function clampNumber(value, min, max, fallback) {
  const parsed = Number.parseInt(String(value ?? ""), 10);
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

function emitProgress(sender, payload) {
  if (!sender || sender.isDestroyed()) {
    return;
  }
  sender.send(PROGRESS_EVENT, payload);
}

async function fileExistsAtPath(filePath) {
  try {
    await stat(filePath);
    return true;
  } catch (error) {
    if (error && error.code === "ENOENT") {
      return false;
    }
    throw error;
  }
}

module.exports = {
  ensureMboxDatabase,
  searchMessages,
  loadMessageById,
  getMessageDateBounds,
  PROGRESS_EVENT
};
