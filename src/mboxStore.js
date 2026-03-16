const { stat, open, rm } = require("fs/promises");
const { createReadStream } = require("fs");
const { setImmediate: waitForImmediate } = require("timers/promises");
const { parseMessageChunk } = require("./mboxParser");
const {
  walkPstMessages,
  buildPstIndexRecord,
  loadPstMessageByDescriptor,
  getPstAttachmentData,
  buildPstEmlBuffer,
  getPstSourcePreview
} = require("./pstConverter");

const SCHEMA_VERSION = "3";
const PARSER_VERSION = "2";
const PROGRESS_EVENT = "mbox-index-progress";
const DB_CACHE = new Map();
const BATCH_SIZE = 250;
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
    // Fall back when node:sqlite is unavailable in Electron.
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

async function ensureMboxDatabase(filePath, sender, options = {}) {
  const sourceStats = await stat(filePath);
  const dbPath = typeof options?.dbPath === "string" && options.dbPath ? options.dbPath : `${filePath}.sqlite`;
  const metaSourcePath = typeof options?.sourcePath === "string" && options.sourcePath ? options.sourcePath : filePath;
  const metaSourceStats = metaSourcePath === filePath ? sourceStats : await stat(metaSourcePath);
  const sourceMtimeMs = Math.trunc(metaSourceStats.mtimeMs);

  emitProgress(sender, {
    phase: "preparing",
    filePath,
    dbPath,
    totalBytes: sourceStats.size,
    bytesRead: 0,
    messagesIndexed: 0
  });

  if (await isReusableDatabase(dbPath, metaSourcePath, metaSourceStats.size, sourceMtimeMs)) {
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
  const statements = createInsertStatements(db);
  let messagesIndexed = 0;
  let transactionOpen = false;
  let lastProgressAt = 0;

  try {
    db.exec("BEGIN");
    transactionOpen = true;

    await streamMboxMessages(
      filePath,
      async ({ rawChunk, byteStart, byteEnd }) => {
        const nextId = messagesIndexed + 1;
        const parsed = parseMessageChunk(rawChunk, {
          index: nextId,
          includeAttachmentData: false,
          includeEmlSource: false,
          includeBodyHtml: false
        });
        if (!parsed) {
          return;
        }

        messagesIndexed = nextId;
        insertIndexRecord(statements, {
          id: nextId,
          sourceKind: "mbox",
          sourceRef: "",
          subject: parsed.subject || "",
          from: parsed.from || "",
          to: parsed.to || "",
          date: parsed.date || "",
          dateTs: parseMessageDateToTimestamp(parsed.date || ""),
          snippet: parsed.snippet || "",
          bodyText: parsed.bodyText || "",
          attachmentNames: (parsed.attachments || [])
            .map((attachment) => attachment.fileName || "")
            .filter(Boolean)
            .join(" "),
          attachments: parsed.attachments || [],
          byteStart,
          byteEnd
        });

        if (messagesIndexed % BATCH_SIZE === 0) {
          db.exec("COMMIT");
          db.exec("BEGIN");
          await waitForImmediate();
        }
      },
      async (progress) => {
        const now = Date.now();
        if (progress.done || now - lastProgressAt >= 200) {
          lastProgressAt = now;
          emitProgress(sender, {
            phase: "indexing",
            filePath,
            dbPath,
            totalBytes: sourceStats.size,
            bytesRead: progress.bytesRead,
            messagesIndexed
          });
          await waitForImmediate();
        }
      }
    );

    db.exec("COMMIT");
    transactionOpen = false;
    finalizeDatabase(db, metaSourcePath, metaSourceStats.size, sourceMtimeMs, messagesIndexed);
  } catch (error) {
    if (transactionOpen) {
      try {
        db.exec("ROLLBACK");
      } catch {
        // Ignore cleanup failures after index errors.
      }
    }
    db.close();
    throw error;
  }

  db.close();
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

async function ensurePstDatabase(filePath, sender, options = {}) {
  const sourceStats = await stat(filePath);
  const dbPath = typeof options?.dbPath === "string" && options.dbPath ? options.dbPath : `${filePath}.sqlite`;
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
  const statements = createInsertStatements(db);
  let messagesIndexed = 0;
  let transactionOpen = false;

  try {
    db.exec("BEGIN");
    transactionOpen = true;

    await walkPstMessages(filePath, {
      onMessage: async (item, index) => {
        const record = buildPstIndexRecord(item, index);
        messagesIndexed = index;
        insertIndexRecord(statements, {
          ...record,
          sourceKind: "pst",
          byteStart: null,
          byteEnd: null
        });

        if (messagesIndexed % BATCH_SIZE === 0) {
          db.exec("COMMIT");
          db.exec("BEGIN");
          await waitForImmediate();
        }
      },
      onProgress: async (progress) => {
        emitProgress(sender, {
          phase: "indexing-pst",
          filePath,
          dbPath,
          totalBytes: sourceStats.size,
          bytesRead: 0,
          messagesIndexed: Number(progress.messagesIndexed) || messagesIndexed
        });
        await waitForImmediate();
      }
    });

    db.exec("COMMIT");
    transactionOpen = false;
    finalizeDatabase(db, filePath, sourceStats.size, sourceMtimeMs, messagesIndexed);
  } catch (error) {
    if (transactionOpen) {
      try {
        db.exec("ROLLBACK");
      } catch {
        // Ignore cleanup failures after index errors.
      }
    }
    db.close();
    throw error;
  }

  db.close();
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
  const row = getMessageRow(dbPath, messageId);
  if (!row) {
    return null;
  }

  const sourcePath = getMetaValue(dbPath, "source_path");
  if (!sourcePath) {
    throw new Error("Indexed database is missing source file metadata.");
  }

  const parsed =
    row.source_kind === "pst"
      ? loadPstMessageByDescriptor(sourcePath, row.source_ref, {
          index: row.id,
          includeAttachmentData: "inline"
        })
      : parseMessageChunk(await readMessageBuffer(sourcePath, row), {
          index: row.id,
          includeAttachmentData: "inline",
          includeEmlSource: false,
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

async function getAttachmentData(dbPath, messageId, attachmentId) {
  const row = getMessageRow(dbPath, messageId);
  if (!row) {
    return null;
  }

  const sourcePath = getMetaValue(dbPath, "source_path");
  if (!sourcePath) {
    return null;
  }

  if (row.source_kind === "pst") {
    return getPstAttachmentData(sourcePath, row.source_ref, attachmentId);
  }

  const parsed = parseMessageChunk(await readMessageBuffer(sourcePath, row), {
    index: row.id,
    includeAttachmentData: true,
    includeEmlSource: false,
    includeBodyHtml: false
  });
  const attachment = (parsed?.attachments || []).find((item) => item.id === attachmentId);
  if (!attachment || !attachment.base64) {
    return null;
  }

  return {
    fileName: attachment.fileName || "attachment.bin",
    contentType: attachment.contentType || "application/octet-stream",
    data: Buffer.from(attachment.base64, "base64")
  };
}

async function getMessageEmlBuffer(dbPath, messageId) {
  const row = getMessageRow(dbPath, messageId);
  if (!row) {
    return null;
  }

  const sourcePath = getMetaValue(dbPath, "source_path");
  if (!sourcePath) {
    return null;
  }

  if (row.source_kind === "pst") {
    return buildPstEmlBuffer(sourcePath, row.source_ref);
  }

  return stripMboxEnvelopeFromBuffer(await readMessageBuffer(sourcePath, row));
}

async function getMessageSourcePreview(dbPath, messageId) {
  const row = getMessageRow(dbPath, messageId);
  if (!row) {
    return "";
  }

  const sourcePath = getMetaValue(dbPath, "source_path");
  if (!sourcePath) {
    return "";
  }

  if (row.source_kind === "pst") {
    return getPstSourcePreview(sourcePath, row.source_ref);
  }

  return stripMboxEnvelopeFromBuffer(await readMessageBuffer(sourcePath, row)).toString("utf8");
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
    await onMessage({
      rawChunk: Buffer.concat(currentLines, currentLength),
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
    const valid =
      readMeta("schema_version") === SCHEMA_VERSION &&
      readMeta("parser_version") === PARSER_VERSION &&
      readMeta("source_path") === sourcePath &&
      Number.parseInt(readMeta("source_size"), 10) === sourceSize &&
      Number.parseInt(readMeta("source_mtime_ms"), 10) === sourceMtimeMs;

    if (!valid) {
      closeDatabase(dbPath);
    }

    return valid;
  } catch {
    closeDatabase(dbPath);
    return false;
  }
}

async function getReusableDatabaseInfo(dbPath, sourcePath) {
  try {
    const sourceStats = await stat(sourcePath);
    const sourceMtimeMs = Math.trunc(sourceStats.mtimeMs);
    const reusable = await isReusableDatabase(dbPath, sourcePath, sourceStats.size, sourceMtimeMs);
    if (!reusable) {
      return null;
    }

    return {
      dbPath,
      totalMessages: countMessages(dbPath),
      sourceSize: sourceStats.size,
      sourceMtimeMs
    };
  } catch {
    return null;
  }
}

function createInsertStatements(db) {
  return {
    insertMessage: db.prepare(`
      INSERT INTO messages (
        id,
        source_kind,
        source_ref,
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
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `),
    insertAttachment: db.prepare(`
      INSERT INTO attachments (
        message_id,
        file_name,
        content_type,
        size,
        is_inline,
        content_id
      ) VALUES (?, ?, ?, ?, ?, ?)
    `),
    insertFts: db.prepare(`
      INSERT INTO message_fts (
        rowid,
        subject,
        sender,
        recipient,
        snippet,
        body_text,
        attachment_names
      ) VALUES (?, ?, ?, ?, ?, ?, ?)
    `)
  };
}

function insertIndexRecord(statements, record) {
  statements.insertMessage.run(
    record.id,
    record.sourceKind,
    record.sourceRef || "",
    record.subject || "",
    record.from || "",
    record.to || "",
    record.date || "",
    record.dateTs ?? null,
    record.snippet || "",
    record.bodyText || "",
    record.attachmentNames || "",
    record.byteStart ?? null,
    record.byteEnd ?? null
  );

  statements.insertFts.run(
    record.id,
    record.subject || "",
    record.from || "",
    record.to || "",
    record.snippet || "",
    record.bodyText || "",
    record.attachmentNames || ""
  );

  for (const attachment of record.attachments || []) {
    statements.insertAttachment.run(
      record.id,
      attachment.fileName || "",
      attachment.contentType || "",
      attachment.size ?? null,
      attachment.isInline ? 1 : 0,
      attachment.contentId || ""
    );
  }
}

function finalizeDatabase(db, sourcePath, sourceSize, sourceMtimeMs, totalMessages) {
  db.exec("ANALYZE");
  db.exec("INSERT INTO message_fts(message_fts) VALUES ('optimize')");
  writeMeta(db, sourcePath, sourceSize, sourceMtimeMs, totalMessages);
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
      source_kind TEXT NOT NULL DEFAULT 'mbox',
      source_ref TEXT NOT NULL DEFAULT '',
      subject TEXT NOT NULL DEFAULT '',
      sender TEXT NOT NULL DEFAULT '',
      recipient TEXT NOT NULL DEFAULT '',
      date_raw TEXT NOT NULL DEFAULT '',
      date_ts INTEGER,
      snippet TEXT NOT NULL DEFAULT '',
      body_text TEXT NOT NULL DEFAULT '',
      attachment_names TEXT NOT NULL DEFAULT '',
      byte_start INTEGER,
      byte_end INTEGER
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
    getMessage: db.prepare(`
      SELECT id, source_kind, source_ref, subject, sender, recipient, date_raw, snippet, byte_start, byte_end
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

async function readRawRange(filePath, byteStart, byteEnd) {
  const start = Number.parseInt(String(byteStart), 10);
  const end = Number.parseInt(String(byteEnd), 10);
  if (!Number.isInteger(start) || !Number.isInteger(end) || end < start) {
    throw new Error("Invalid message byte offsets in index.");
  }

  const length = end - start;
  if (length === 0) {
    return Buffer.alloc(0);
  }

  const buffer = Buffer.alloc(length);
  const handle = await open(filePath, "r");
  try {
    const { bytesRead } = await handle.read(buffer, 0, length, start);
    return buffer.subarray(0, bytesRead);
  } finally {
    await handle.close();
  }
}

async function readMessageBuffer(sourcePath, row) {
  if (row.source_kind !== "mbox") {
    return Buffer.alloc(0);
  }
  return readRawRange(sourcePath, row.byte_start, row.byte_end);
}

function stripMboxEnvelopeFromBuffer(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    return Buffer.alloc(0);
  }

  if (
    buffer.length >= 5 &&
    buffer[0] === 0x46 &&
    buffer[1] === 0x72 &&
    buffer[2] === 0x6f &&
    buffer[3] === 0x6d &&
    buffer[4] === 0x20
  ) {
    const newlineIndex = buffer.indexOf(0x0a);
    if (newlineIndex !== -1 && newlineIndex + 1 < buffer.length) {
      return buffer.subarray(newlineIndex + 1);
    }
  }

  return buffer;
}

function getMessageRow(dbPath, messageId) {
  const entry = getDatabaseEntry(dbPath);
  const id = Number.parseInt(String(messageId), 10);
  if (!Number.isInteger(id) || id <= 0) {
    return null;
  }

  return entry.getMessage.get(id) || null;
}

function getMetaValue(dbPath, key) {
  return getDatabaseEntry(dbPath).getMeta.get(key)?.value || "";
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
  ensurePstDatabase,
  getReusableDatabaseInfo,
  searchMessages,
  loadMessageById,
  getAttachmentData,
  getMessageEmlBuffer,
  getMessageSourcePreview,
  getMessageDateBounds,
  PROGRESS_EVENT
};
