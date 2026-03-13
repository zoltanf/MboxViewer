const path = require("path");
const { createWriteStream } = require("fs");
const { stat, readFile, writeFile, rm } = require("fs/promises");
const { once } = require("events");

const PST_CONVERTER_VERSION = "2";

function isPstFilePath(filePath) {
  return String(filePath || "").toLowerCase().endsWith(".pst");
}

async function ensurePstConvertedToMbox(pstPath, options = {}) {
  const sourceStats = await stat(pstPath);
  const sourceMtimeMs = Math.trunc(sourceStats.mtimeMs);
  const mboxPath = `${pstPath}.mbox`;
  const metaPath = `${mboxPath}.meta.json`;
  const onProgress = typeof options.onProgress === "function" ? options.onProgress : null;

  const reusable = await isReusableConversion(metaPath, mboxPath, pstPath, sourceStats.size, sourceMtimeMs);
  if (reusable) {
    return {
      mboxPath,
      reused: true,
      messageCount: reusable.messageCount || 0
    };
  }

  let messageCount = 0;
  try {
    if (onProgress) {
      onProgress({ phase: "converting-pst", messagesConverted: 0 });
    }
    messageCount = await convertPstFileToMbox(pstPath, mboxPath, onProgress);
  } catch (error) {
    await safeRemoveFile(mboxPath);
    throw error;
  }

  const meta = {
    converterVersion: PST_CONVERTER_VERSION,
    sourcePath: pstPath,
    sourceSize: sourceStats.size,
    sourceMtimeMs,
    messageCount,
    convertedAt: new Date().toISOString()
  };
  await writeFile(metaPath, JSON.stringify(meta, null, 2), "utf8");

  return {
    mboxPath,
    reused: false,
    messageCount
  };
}

async function isReusableConversion(metaPath, mboxPath, sourcePath, sourceSize, sourceMtimeMs) {
  try {
    const raw = await readFile(metaPath, "utf8");
    const meta = JSON.parse(raw);
    if (!meta || typeof meta !== "object") {
      return null;
    }

    const valid =
      meta.converterVersion === PST_CONVERTER_VERSION &&
      meta.sourcePath === sourcePath &&
      Number(meta.sourceSize) === Number(sourceSize) &&
      Number(meta.sourceMtimeMs) === Number(sourceMtimeMs);

    if (!valid) {
      return null;
    }

    await stat(mboxPath);

    return {
      messageCount: Number(meta.messageCount) || 0
    };
  } catch {
    return null;
  }
}

async function convertPstFileToMbox(pstPath, mboxPath, onProgress = null) {
  const { PSTFile } = require("pst-extractor");
  const pstFile = new PSTFile(pstPath);
  const output = createWriteStream(mboxPath, { encoding: "utf8" });

  let streamError = null;
  output.on("error", (error) => {
    streamError = error;
  });

  let messageCount = 0;
  let lastProgressAt = 0;

  try {
    const rootFolder = pstFile.getRootFolder();

    await walkFolderTree(rootFolder, async (item) => {
      if (!isMailLikePstItem(item)) {
        return;
      }

      messageCount += 1;
      const raw = buildMboxMessage(item, messageCount);
      await writeChunk(output, raw);
      if (onProgress) {
        const now = Date.now();
        if (messageCount === 1 || now - lastProgressAt >= 150) {
          lastProgressAt = now;
          onProgress({ phase: "converting-pst", messagesConverted: messageCount });
        }
      }
      if (streamError) {
        throw streamError;
      }
    });
    await endStream(output);
  } catch (error) {
    output.destroy();
    throw error;
  } finally {
    try {
      pstFile.close();
    } catch {
      // Ignore close errors from converter cleanup.
    }
  }

  if (streamError) {
    throw streamError;
  }

  return messageCount;
}

async function walkFolderTree(folder, onMessage) {
  if (!folder) {
    return;
  }

  if (folder.hasSubfolders) {
    const subFolders = folder.getSubFolders() || [];
    for (const child of subFolders) {
      await walkFolderTree(child, onMessage);
    }
  }

  if ((Number(folder.contentCount) || 0) <= 0) {
    return;
  }

  if (typeof folder.moveChildCursorTo === "function") {
    folder.moveChildCursorTo(0);
  }

  let child = folder.getNextChild();
  while (child != null) {
    await onMessage(child);
    child = folder.getNextChild();
  }
}

function isMailLikePstItem(item) {
  if (!item || typeof item !== "object") {
    return false;
  }

  const messageClass = String(item.messageClass || "").trim().toUpperCase();
  if (messageClass) {
    if (!messageClass.startsWith("IPM.")) {
      return false;
    }

    const excluded = [
      "IPM.CONTACT",
      "IPM.APPOINTMENT",
      "IPM.TASK",
      "IPM.STICKYNOTE",
      "IPM.JOURNAL",
      "IPM.ACTIVITY"
    ];
    if (excluded.some((prefix) => messageClass.startsWith(prefix))) {
      return false;
    }
  }

  const hasBody = Boolean(String(item.body || "").trim() || String(item.bodyHTML || "").trim());
  const hasSubject = Boolean(String(item.subject || "").trim());
  return hasBody || hasSubject;
}

function buildMboxMessage(message, index) {
  const submittedAt = message.clientSubmitTime instanceof Date && Number.isFinite(message.clientSubmitTime.getTime())
    ? message.clientSubmitTime
    : new Date(0);

  const senderEmail = resolveEnvelopeSender(message);
  const fromLine = `From ${senderEmail} ${formatEnvelopeDate(submittedAt)}`;

  const fromHeader = buildFromHeader(message);
  const toHeader = normalizeHeaderValue(message.displayTo || "");
  const ccHeader = normalizeHeaderValue(message.displayCC || "");
  const subject = normalizeHeaderValue(message.subject || "(No Subject)");
  const dateHeader = submittedAt.getTime() > 0 ? submittedAt.toUTCString() : "";

  const htmlBody = normalizeBodyValue(message.bodyHTML || "");
  const textBody = normalizeBodyValue(message.body || "");
  const useHtml = Boolean(htmlBody);
  const primaryBody = useHtml ? htmlBody : textBody;
  const attachments = extractMessageAttachments(message);

  const headers = [];
  headers.push(`Subject: ${subject}`);
  if (fromHeader) {
    headers.push(`From: ${fromHeader}`);
  }
  if (toHeader) {
    headers.push(`To: ${toHeader}`);
  }
  if (ccHeader) {
    headers.push(`Cc: ${ccHeader}`);
  }
  if (dateHeader) {
    headers.push(`Date: ${dateHeader}`);
  }
  headers.push(`Message-ID: <pst-${index}@mboxviewer.local>`);
  headers.push("MIME-Version: 1.0");

  if (attachments.length === 0) {
    headers.push(`Content-Type: ${useHtml ? "text/html" : "text/plain"}; charset=utf-8`);
    headers.push("Content-Transfer-Encoding: 8bit");
    const escapedBody = escapeMboxBody(primaryBody);
    return `${fromLine}\n${headers.join("\n")}\n\n${escapedBody}\n\n`;
  }

  const boundary = buildMimeBoundary(index, submittedAt.getTime());
  headers.push(`Content-Type: multipart/mixed; boundary=\"${boundary}\"`);
  const multipartBody = buildMultipartMessageBody(boundary, useHtml, primaryBody, attachments);
  return `${fromLine}\n${headers.join("\n")}\n\n${multipartBody}\n\n`;
}

function extractMessageAttachments(message) {
  const attachmentCount = Number(message?.numberOfAttachments) || 0;
  if (attachmentCount <= 0) {
    return [];
  }

  const attachments = [];

  for (let index = 0; index < attachmentCount; index += 1) {
    let attachment = null;
    try {
      attachment = message.getAttachment(index);
    } catch {
      continue;
    }

    if (!attachment) {
      continue;
    }

    const fileBuffer = readAttachmentBuffer(attachment);
    if (!fileBuffer || fileBuffer.length === 0) {
      continue;
    }

    const fileName = resolveAttachmentFileName(attachment, index);
    const contentType = resolveAttachmentContentType(attachment, fileName);
    const contentId = normalizeContentId(attachment.contentId || "");
    const isInline = Boolean(contentId) && !attachment.isAttachmentInvisibleInHtml;

    attachments.push({
      fileName,
      contentType,
      contentId,
      isInline,
      base64: toBase64Lines(fileBuffer)
    });
  }

  return attachments;
}

function readAttachmentBuffer(attachment) {
  try {
    const stream = attachment.fileInputStream;
    if (!stream) {
      return null;
    }

    const chunks = [];
    const chunkSize = 8192;
    let bytesRead = 0;

    for (;;) {
      const chunk = Buffer.alloc(chunkSize);
      bytesRead = stream.readBlock(chunk);
      if (bytesRead <= 0) {
        break;
      }
      chunks.push(bytesRead === chunk.length ? chunk : chunk.slice(0, bytesRead));
    }

    return chunks.length > 0 ? Buffer.concat(chunks) : Buffer.alloc(0);
  } catch {
    return null;
  }
}

function resolveAttachmentFileName(attachment, index) {
  const candidates = [
    attachment.longFilename,
    attachment.filename,
    attachment.pathname ? path.basename(String(attachment.pathname)) : "",
    attachment.longPathname ? path.basename(String(attachment.longPathname)) : "",
    `attachment-${index + 1}.bin`
  ];

  for (const candidate of candidates) {
    const sanitized = sanitizeAttachmentFileName(candidate);
    if (sanitized) {
      return sanitized;
    }
  }

  return `attachment-${index + 1}.bin`;
}

function sanitizeAttachmentFileName(input) {
  return String(input || "")
    .replace(/[\x00-\x1F]/g, "")
    .replace(/[\\/:"*?<>|]+/g, "_")
    .replace(/\s+/g, " ")
    .trim();
}

function resolveAttachmentContentType(attachment, fileName) {
  const mimeTag = String(attachment.mimeTag || "")
    .trim()
    .toLowerCase();
  if (/^[a-z0-9!#$&^_.+-]+\/[a-z0-9!#$&^_.+-]+$/.test(mimeTag)) {
    return mimeTag;
  }
  return mimeFromExtension(fileName);
}

function mimeFromExtension(fileName) {
  const extension = path.extname(String(fileName || "")).toLowerCase();
  const map = {
    ".txt": "text/plain",
    ".html": "text/html",
    ".htm": "text/html",
    ".csv": "text/csv",
    ".json": "application/json",
    ".xml": "application/xml",
    ".pdf": "application/pdf",
    ".zip": "application/zip",
    ".gz": "application/gzip",
    ".doc": "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xls": "application/vnd.ms-excel",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".ppt": "application/vnd.ms-powerpoint",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".eml": "message/rfc822",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".svg": "image/svg+xml",
    ".bmp": "image/bmp",
    ".ico": "image/x-icon",
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".mp4": "video/mp4"
  };
  return map[extension] || "application/octet-stream";
}

function normalizeContentId(value) {
  return String(value || "")
    .replace(/[\r\n]/g, "")
    .trim()
    .replace(/^<|>$/g, "");
}

function buildMimeBoundary(index, timestampMs) {
  const stamp = Number.isFinite(timestampMs) && timestampMs > 0 ? Math.trunc(timestampMs) : index;
  return `----mboxviewer-pst-${index}-${stamp.toString(16)}`;
}

function buildMultipartMessageBody(boundary, useHtml, primaryBody, attachments) {
  const parts = [];

  const bodyBuffer = Buffer.from(primaryBody || "", "utf8");
  parts.push(
    [
      `--${boundary}`,
      `Content-Type: ${useHtml ? "text/html" : "text/plain"}; charset=utf-8`,
      "Content-Transfer-Encoding: base64",
      "Content-Disposition: inline",
      "",
      toBase64Lines(bodyBuffer)
    ].join("\n")
  );

  for (const attachment of attachments) {
    const encodedName = escapeMimeHeaderParam(attachment.fileName || "attachment.bin");
    const headers = [
      `--${boundary}`,
      `Content-Type: ${attachment.contentType || "application/octet-stream"}; name=\"${encodedName}\"`,
      "Content-Transfer-Encoding: base64",
      `Content-Disposition: ${attachment.isInline ? "inline" : "attachment"}; filename=\"${encodedName}\"`
    ];
    if (attachment.contentId) {
      headers.push(`Content-ID: <${normalizeContentId(attachment.contentId)}>`);
    }
    headers.push("", attachment.base64 || "");
    parts.push(headers.join("\n"));
  }

  parts.push(`--${boundary}--`);
  return parts.join("\n");
}

function escapeMimeHeaderParam(value) {
  return String(value || "")
    .replace(/[\r\n]+/g, " ")
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"')
    .trim();
}

function toBase64Lines(buffer) {
  const value = Buffer.isBuffer(buffer) ? buffer.toString("base64") : Buffer.from(String(buffer || ""), "utf8").toString("base64");
  if (!value) {
    return "";
  }
  const lines = [];
  for (let index = 0; index < value.length; index += 76) {
    lines.push(value.slice(index, index + 76));
  }
  return lines.join("\n");
}

function resolveEnvelopeSender(message) {
  const senderCandidates = [
    message.senderEmailAddress,
    extractEmailAddress(String(message.senderName || "")),
    "unknown@pst.local"
  ];

  for (const candidate of senderCandidates) {
    const normalized = sanitizeMailboxToken(candidate);
    if (normalized) {
      return normalized;
    }
  }

  return "unknown@pst.local";
}

function buildFromHeader(message) {
  const name = normalizeHeaderValue(message.senderName || "");
  const email = sanitizeMailboxToken(message.senderEmailAddress);
  if (name && email) {
    return `${name} <${email}>`;
  }
  if (email) {
    return email;
  }
  return name;
}

function extractEmailAddress(input) {
  const match = String(input || "").match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0] : "";
}

function sanitizeMailboxToken(input) {
  const value = String(input || "").trim();
  if (!value) {
    return "";
  }

  const cleaned = value.replace(/[\s<>\"']/g, "");
  return /.+@.+/.test(cleaned) ? cleaned : "";
}

function normalizeHeaderValue(value) {
  return String(value || "")
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeBodyValue(value) {
  return String(value || "")
    .replace(/\u0000/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
}

function escapeMboxBody(body) {
  return String(body || "")
    .split("\n")
    .map((line) => (line.startsWith("From ") ? `>${line}` : line))
    .join("\n");
}

function formatEnvelopeDate(date) {
  const parsed = date instanceof Date && Number.isFinite(date.getTime()) ? date : new Date(0);
  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const dayName = dayNames[parsed.getDay()];
  const month = monthNames[parsed.getMonth()];
  const day = String(parsed.getDate()).padStart(2, " ");
  const hh = String(parsed.getHours()).padStart(2, "0");
  const mm = String(parsed.getMinutes()).padStart(2, "0");
  const ss = String(parsed.getSeconds()).padStart(2, "0");
  const year = parsed.getFullYear();
  return `${dayName} ${month} ${day} ${hh}:${mm}:${ss} ${year}`;
}

async function writeChunk(stream, text) {
  if (stream.write(text)) {
    return;
  }
  await once(stream, "drain");
}

async function endStream(stream) {
  await new Promise((resolve, reject) => {
    stream.on("error", reject);
    stream.end(resolve);
  });
}

async function safeRemoveFile(filePath) {
  try {
    await rm(filePath);
  } catch (error) {
    if (!error || error.code !== "ENOENT") {
      throw error;
    }
  }
}

module.exports = {
  isPstFilePath,
  ensurePstConvertedToMbox
};
