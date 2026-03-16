const path = require("path");
const Long = require("long");
const { PSTFile } = require("pst-extractor");
const { PSTUtil } = require("pst-extractor/dist/PSTUtil.class");
const { compactWhitespace, plainToHtml, stripHtmlTags } = require("./mboxParser");

const PST_FILE_CACHE = new Map();

function isPstFilePath(filePath) {
  return String(filePath || "").toLowerCase().endsWith(".pst");
}

async function walkPstMessages(pstPath, options = {}) {
  const onMessage = typeof options.onMessage === "function" ? options.onMessage : null;
  const onProgress = typeof options.onProgress === "function" ? options.onProgress : null;
  const pstFile = getCachedPstFile(pstPath);
  const rootFolder = pstFile.getRootFolder();
  let messageCount = 0;
  let lastProgressAt = 0;

  await walkFolderTree(rootFolder, async (item) => {
    if (!isMailLikePstItem(item)) {
      return;
    }

    messageCount += 1;
    if (onMessage) {
      await onMessage(item, messageCount);
    }

    if (onProgress) {
      const now = Date.now();
      if (messageCount === 1 || now - lastProgressAt >= 150) {
        lastProgressAt = now;
        onProgress({ phase: "indexing-pst", messagesIndexed: messageCount });
      }
    }
  });

  if (onProgress) {
    onProgress({ phase: "indexing-pst", messagesIndexed: messageCount, done: true });
  }

  return { messageCount };
}

function buildPstIndexRecord(message, index) {
  const submittedAt = getMessageSubmitDate(message);
  const from = buildFromHeader(message);
  const to = normalizeHeaderValue(message.displayTo || "");
  const date = submittedAt ? submittedAt.toUTCString() : "";
  const htmlBody = normalizeBodyValue(message.bodyHTML || "");
  const textBody = normalizeBodyValue(message.body || "") || stripHtmlTags(htmlBody);
  const attachments = extractMessageAttachments(message, { includeAttachmentData: "none" });

  return {
    id: index,
    sourceRef: String(message.descriptorNodeId),
    subject: normalizeHeaderValue(message.subject || "(No Subject)"),
    from,
    to,
    date,
    dateTs: submittedAt ? submittedAt.getTime() : null,
    snippet: compactWhitespace(textBody).slice(0, 180),
    bodyText: textBody,
    attachmentNames: attachments.map((attachment) => attachment.fileName || "").filter(Boolean).join(" "),
    attachments
  };
}

function loadPstMessageByDescriptor(pstPath, descriptorId, options = {}) {
  const message = getPstMessageByDescriptor(pstPath, descriptorId);
  if (!message || !isMailLikePstItem(message)) {
    return null;
  }

  const index = Number.parseInt(String(options.index || 0), 10) || 0;
  const submittedAt = getMessageSubmitDate(message);
  const subject = normalizeHeaderValue(message.subject || "(No Subject)");
  const from = buildFromHeader(message);
  const to = normalizeHeaderValue(message.displayTo || "");
  const date = submittedAt ? submittedAt.toUTCString() : "";
  const bodyHtmlRaw = normalizeBodyValue(message.bodyHTML || "");
  const bodyText = normalizeBodyValue(message.body || "") || stripHtmlTags(bodyHtmlRaw);
  const attachments = extractMessageAttachments(message, {
    includeAttachmentData: options.includeAttachmentData === "all" ? "all" : "inline"
  });

  return {
    id: index || String(descriptorId),
    subject,
    from,
    to,
    date,
    snippet: compactWhitespace(bodyText).slice(0, 180),
    bodyText,
    bodyHtml: bodyHtmlRaw || plainToHtml(bodyText),
    attachments
  };
}

function getPstAttachmentData(pstPath, descriptorId, attachmentId) {
  const message = getPstMessageByDescriptor(pstPath, descriptorId);
  if (!message || !isMailLikePstItem(message)) {
    return null;
  }

  const attachmentIndex = parseAttachmentId(attachmentId);
  if (attachmentIndex === null) {
    return null;
  }

  let attachment = null;
  try {
    attachment = message.getAttachment(attachmentIndex);
  } catch {
    return null;
  }

  if (!attachment) {
    return null;
  }

  const fileBuffer = readAttachmentBuffer(attachment);
  if (!fileBuffer) {
    return null;
  }

  const fileName = resolveAttachmentFileName(attachment, attachmentIndex);
  const contentType = resolveAttachmentContentType(attachment, fileName);
  return {
    fileName,
    contentType,
    data: fileBuffer
  };
}

function buildPstEmlBuffer(pstPath, descriptorId) {
  const message = getPstMessageByDescriptor(pstPath, descriptorId);
  if (!message || !isMailLikePstItem(message)) {
    return null;
  }

  return Buffer.from(buildRfc822Message(message, String(descriptorId)), "utf8");
}

function getPstSourcePreview(pstPath, descriptorId) {
  const buffer = buildPstEmlBuffer(pstPath, descriptorId);
  return buffer ? buffer.toString("utf8") : "";
}

function getPstMessageByDescriptor(pstPath, descriptorId) {
  const pstFile = getCachedPstFile(pstPath);
  return PSTUtil.detectAndLoadPSTObject(pstFile, Long.fromString(String(descriptorId)));
}

function getCachedPstFile(pstPath) {
  const normalizedPath = path.resolve(String(pstPath || ""));
  const existing = PST_FILE_CACHE.get(normalizedPath);
  if (existing) {
    return existing;
  }

  const pstFile = new PSTFile(normalizedPath);
  PST_FILE_CACHE.set(normalizedPath, pstFile);
  return pstFile;
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

function buildRfc822Message(message, messageIdToken) {
  const submittedAt = getMessageSubmitDate(message);
  const fromHeader = buildFromHeader(message);
  const toHeader = normalizeHeaderValue(message.displayTo || "");
  const ccHeader = normalizeHeaderValue(message.displayCC || "");
  const subject = normalizeHeaderValue(message.subject || "(No Subject)");
  const dateHeader = submittedAt ? submittedAt.toUTCString() : "";
  const htmlBody = normalizeBodyValue(message.bodyHTML || "");
  const textBody = normalizeBodyValue(message.body || "");
  const useHtml = Boolean(htmlBody);
  const primaryBody = useHtml ? htmlBody : textBody;
  const attachments = extractMessageAttachments(message, { includeAttachmentData: "all" });
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
  headers.push(`Message-ID: <pst-${messageIdToken}@mboxviewer.local>`);
  headers.push("MIME-Version: 1.0");

  if (attachments.length === 0) {
    headers.push(`Content-Type: ${useHtml ? "text/html" : "text/plain"}; charset=utf-8`);
    headers.push("Content-Transfer-Encoding: 8bit");
    return `${headers.join("\n")}\n\n${primaryBody}\n`;
  }

  const boundary = buildMimeBoundary(messageIdToken, submittedAt ? submittedAt.getTime() : 0);
  headers.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
  return `${headers.join("\n")}\n\n${buildMultipartMessageBody(boundary, useHtml, primaryBody, attachments)}\n`;
}

function extractMessageAttachments(message, options = {}) {
  const includeAttachmentData = options.includeAttachmentData || "none";
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

    const fileName = resolveAttachmentFileName(attachment, index);
    const contentType = resolveAttachmentContentType(attachment, fileName);
    const contentId = normalizeContentId(attachment.contentId || "");
    const isInline = Boolean(contentId) && !attachment.isAttachmentInvisibleInHtml;
    const shouldReadData =
      includeAttachmentData === "all" ||
      (includeAttachmentData === "inline" && isInline);
    const fileBuffer = shouldReadData ? readAttachmentBuffer(attachment) : null;

    attachments.push({
      id: buildAttachmentId(index),
      fileName,
      contentType,
      size: Number(attachment.filesize || attachment.size) || (fileBuffer ? fileBuffer.length : null),
      isInline,
      contentId,
      base64: fileBuffer ? fileBuffer.toString("base64") : ""
    });
  }

  return attachments;
}

function parseAttachmentId(attachmentId) {
  const match = String(attachmentId || "").match(/^pst-att-(\d+)$/);
  if (!match) {
    return null;
  }

  const parsed = Number.parseInt(match[1], 10);
  return Number.isInteger(parsed) ? parsed : null;
}

function buildAttachmentId(index) {
  return `pst-att-${index}`;
}

function readAttachmentBuffer(attachment) {
  try {
    const stream = attachment.fileInputStream;
    if (!stream) {
      return null;
    }

    const chunks = [];
    const chunkSize = 8192;
    for (;;) {
      const chunk = Buffer.alloc(chunkSize);
      const bytesRead = stream.readBlock(chunk);
      if (bytesRead <= 0) {
        break;
      }
      chunks.push(bytesRead === chunk.length ? chunk : chunk.subarray(0, bytesRead));
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
  const mimeTag = String(attachment.mimeTag || "").trim().toLowerCase();
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

function buildMimeBoundary(messageIdToken, timestampMs) {
  const stamp = Number.isFinite(timestampMs) && timestampMs > 0 ? Math.trunc(timestampMs) : Number(messageIdToken) || 0;
  return `----mboxviewer-pst-${messageIdToken}-${stamp.toString(16)}`;
}

function buildMultipartMessageBody(boundary, useHtml, primaryBody, attachments) {
  const parts = [];
  parts.push(
    [
      `--${boundary}`,
      `Content-Type: ${useHtml ? "text/html" : "text/plain"}; charset=utf-8`,
      "Content-Transfer-Encoding: base64",
      "Content-Disposition: inline",
      "",
      toBase64Lines(Buffer.from(primaryBody || "", "utf8"))
    ].join("\n")
  );

  for (const attachment of attachments) {
    const encodedName = escapeMimeHeaderParam(attachment.fileName || "attachment.bin");
    const headers = [
      `--${boundary}`,
      `Content-Type: ${attachment.contentType || "application/octet-stream"}; name="${encodedName}"`,
      "Content-Transfer-Encoding: base64",
      `Content-Disposition: ${attachment.isInline ? "inline" : "attachment"}; filename="${encodedName}"`
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

function getMessageSubmitDate(message) {
  const submittedAt = message.clientSubmitTime;
  if (!(submittedAt instanceof Date) || !Number.isFinite(submittedAt.getTime()) || submittedAt.getTime() <= 0) {
    return null;
  }
  return submittedAt;
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

module.exports = {
  isPstFilePath,
  walkPstMessages,
  buildPstIndexRecord,
  loadPstMessageByDescriptor,
  getPstAttachmentData,
  buildPstEmlBuffer,
  getPstSourcePreview
};
