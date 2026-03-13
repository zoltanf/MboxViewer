function parseMbox(raw, options = {}) {
  const lines = raw.replace(/\r\n/g, "\n").split("\n");
  const chunks = [];
  let current = [];

  for (const line of lines) {
    if (line.startsWith("From ") && current.length > 0) {
      chunks.push(current.join("\n"));
      current = [line];
      continue;
    }
    current.push(line);
  }

  if (current.length > 0) {
    chunks.push(current.join("\n"));
  }

  return chunks
    .map((chunk, index) => parseMessageChunk(chunk, { ...options, index }))
    .filter(Boolean);
}

function parseMessageChunk(chunk, options = {}) {
  const normalizedOptions = {
    includeAttachmentData: options.includeAttachmentData !== false,
    includeEmlSource: options.includeEmlSource !== false,
    includeBodyHtml: options.includeBodyHtml !== false
  };
  const index = Number.isInteger(options.index) ? options.index : 0;
  const lines = chunk.split("\n");
  if (lines.length === 0) {
    return null;
  }

  if (lines[0].startsWith("From ")) {
    lines.shift();
  }
  const emlSource = normalizedOptions.includeEmlSource ? lines.join("\n") : "";

  const separatorIndex = lines.findIndex((line) => line.trim() === "");
  const headerLines = separatorIndex === -1 ? lines : lines.slice(0, separatorIndex);
  const bodyLines = separatorIndex === -1 ? [] : lines.slice(separatorIndex + 1);

  const headers = parseHeaders(headerLines);
  const bodyRaw = bodyLines.join("\n");
  const parsed = parseEntity(headers, bodyRaw, { attachmentCounter: 0 }, normalizedOptions);

  const subject = decodeMimeWords(headers.subject || "(No Subject)");
  const from = decodeMimeWords(headers.from || "");
  const to = decodeMimeWords(headers.to || "");
  const date = headers.date || "";

  const bodyHtml = normalizedOptions.includeBodyHtml ? parsed.html || plainToHtml(parsed.text || "") : "";
  const bodyText = parsed.text || stripHtmlTags(bodyHtml);
  const snippet = compactWhitespace(bodyText).slice(0, 180);

  const attachments = parsed.attachments.map((attachment, attIndex) => ({
    id: `${index}-att-${attIndex}`,
    fileName: attachment.fileName,
    contentType: attachment.contentType,
    size: attachment.size,
    isInline: attachment.isInline,
    contentId: attachment.contentId,
    base64: normalizedOptions.includeAttachmentData ? attachment.base64 : ""
  }));

  return {
    id: `${index}-${hash(`${subject}|${from}|${date}`)}`,
    subject,
    from,
    to,
    date,
    snippet,
    bodyText,
    bodyHtml,
    attachments,
    emlSource: normalizedOptions.includeEmlSource ? emlSource : ""
  };
}

function parseEntity(headers, bodyRaw, state, options) {
  const contentTypeRaw = headers["content-type"] || "text/plain";
  const contentType = parseHeaderValue(contentTypeRaw);
  const dispositionRaw = headers["content-disposition"] || "";
  const disposition = parseHeaderValue(dispositionRaw);
  const transferEncoding = (headers["content-transfer-encoding"] || "").toLowerCase();

  if (contentType.mediaType.startsWith("multipart/")) {
    const boundary = contentType.params.boundary || extractBoundary(contentTypeRaw);
    if (!boundary) {
      const fallbackText = decodeBuffer(decodeTransferToBuffer(bodyRaw, transferEncoding), contentType.params.charset);
      return { text: fallbackText, html: "", attachments: [] };
    }

    const parts = splitMultipart(bodyRaw, boundary);
    const parsedParts = parts
      .map((partRaw) => parseRawPart(partRaw))
      .filter(Boolean)
      .map((part) => parseEntity(part.headers, part.bodyRaw, state, options));

    const attachments = [];
    let text = "";
    let html = "";

    if (contentType.mediaType === "multipart/alternative") {
      for (const part of parsedParts) {
        attachments.push(...part.attachments);
        if (part.text) {
          text = part.text;
        }
        if (part.html) {
          html = part.html;
        }
      }
    } else {
      for (const part of parsedParts) {
        attachments.push(...part.attachments);
        if (!text && part.text) {
          text = part.text;
        }
        if (!html && part.html) {
          html = part.html;
        }
      }
    }

    if (!text && html) {
      text = stripHtmlTags(html);
    }
    if (!html && text && options.includeBodyHtml) {
      html = plainToHtml(text);
    }

    return { text, html, attachments };
  }

  const mediaType = contentType.mediaType;

  const fileName = extractFileName(contentType.params, disposition.params);
  const contentId = normalizeContentId(headers["content-id"] || "");
  const isTextPlain = mediaType === "text/plain";
  const isTextHtml = mediaType === "text/html";
  const isAttachment = disposition.mediaType === "attachment" || Boolean(fileName) || (!isTextPlain && !isTextHtml);
  const isInline = disposition.mediaType === "inline" || (mediaType.startsWith("image/") && Boolean(contentId));

  if ((isTextPlain || isTextHtml) && !isAttachment) {
    const textBuffer = decodeTransferToBuffer(bodyRaw, transferEncoding);
    const decodedText = decodeBuffer(textBuffer, contentType.params.charset);
    if (isTextHtml) {
      return { text: stripHtmlTags(decodedText), html: options.includeBodyHtml ? decodedText : "", attachments: [] };
    }
    return { text: decodedText, html: options.includeBodyHtml ? plainToHtml(decodedText) : "", attachments: [] };
  }

  const resolvedName = fileName || defaultAttachmentName(mediaType, state);
  const textBuffer = options.includeAttachmentData ? decodeTransferToBuffer(bodyRaw, transferEncoding) : null;
  return {
    text: "",
    html: "",
    attachments: [
      {
        fileName: resolvedName,
        contentType: mediaType || "application/octet-stream",
        size: textBuffer ? textBuffer.length : null,
        isInline,
        contentId,
        base64: textBuffer ? textBuffer.toString("base64") : ""
      }
    ]
  };
}

function parseRawPart(rawPart) {
  const lines = rawPart.replace(/\r\n/g, "\n").split("\n");
  const separatorIndex = lines.findIndex((line) => line.trim() === "");
  const headerLines = separatorIndex === -1 ? lines : lines.slice(0, separatorIndex);
  const bodyLines = separatorIndex === -1 ? [] : lines.slice(separatorIndex + 1);
  const headers = parseHeaders(headerLines);
  return { headers, bodyRaw: bodyLines.join("\n") };
}

function parseHeaders(lines) {
  const unfolded = [];

  for (const line of lines) {
    if (/^\s/.test(line) && unfolded.length > 0) {
      unfolded[unfolded.length - 1] += ` ${line.trim()}`;
    } else {
      unfolded.push(line);
    }
  }

  const headers = {};
  for (const line of unfolded) {
    const idx = line.indexOf(":");
    if (idx === -1) {
      continue;
    }
    const key = line.slice(0, idx).trim().toLowerCase();
    const value = line.slice(idx + 1).trim();
    headers[key] = value;
  }

  return headers;
}

function parseHeaderValue(rawValue) {
  const segments = splitHeaderSegments(rawValue || "");
  const mediaType = (segments.shift() || "").trim().toLowerCase();
  const params = {};

  for (const segment of segments) {
    const idx = segment.indexOf("=");
    if (idx === -1) {
      continue;
    }
    const key = segment.slice(0, idx).trim().toLowerCase();
    let value = segment.slice(idx + 1).trim();
    if (value.startsWith('"') && value.endsWith('"')) {
      value = value.slice(1, -1);
    }
    params[key] = value;
  }

  return { mediaType, params };
}

function splitHeaderSegments(value) {
  const segments = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < value.length; i += 1) {
    const char = value[i];
    if (char === '"') {
      inQuotes = !inQuotes;
      current += char;
      continue;
    }
    if (char === ";" && !inQuotes) {
      segments.push(current);
      current = "";
      continue;
    }
    current += char;
  }

  if (current) {
    segments.push(current);
  }

  return segments;
}

function extractBoundary(contentTypeHeader) {
  const match = contentTypeHeader.match(/boundary\*?="?([^";]+)"?/i);
  if (!match) {
    return "";
  }

  if (match[0].toLowerCase().includes("boundary*=")) {
    return decodeRfc2231Value(match[1]);
  }

  return match[1];
}

function splitMultipart(bodyRaw, boundary) {
  const boundaryLine = `--${boundary}`;
  const endBoundaryLine = `--${boundary}--`;
  const lines = bodyRaw.replace(/\r\n/g, "\n").split("\n");

  const parts = [];
  let current = [];
  let inPart = false;

  for (const line of lines) {
    if (line === boundaryLine) {
      if (inPart && current.length > 0) {
        parts.push(current.join("\n"));
      }
      inPart = true;
      current = [];
      continue;
    }

    if (line === endBoundaryLine) {
      if (inPart && current.length > 0) {
        parts.push(current.join("\n"));
      }
      break;
    }

    if (inPart) {
      current.push(line);
    }
  }

  return parts;
}

function extractFileName(contentTypeParams, dispositionParams) {
  const rawName =
    dispositionParams["filename*"] ||
    contentTypeParams["name*"] ||
    dispositionParams.filename ||
    contentTypeParams.name ||
    "";

  if (!rawName) {
    return "";
  }

  const decoded = rawName.includes("''") ? decodeRfc2231Value(rawName) : rawName;
  return decodeMimeWords(decoded).trim();
}

function decodeRfc2231Value(input) {
  const clean = String(input || "").trim();
  const match = clean.match(/^([^']*)'[^']*'(.*)$/);
  if (!match) {
    try {
      return decodeURIComponent(clean);
    } catch {
      return clean;
    }
  }

  const charset = (match[1] || "utf-8").toLowerCase();
  const encoded = match[2] || "";
  const bytes = decodePercentToBuffer(encoded);
  return decodeBuffer(bytes, charset);
}

function decodePercentToBuffer(input) {
  const bytes = [];

  for (let i = 0; i < input.length; i += 1) {
    if (input[i] === "%" && /[A-Fa-f0-9]{2}/.test(input.slice(i + 1, i + 3))) {
      bytes.push(parseInt(input.slice(i + 1, i + 3), 16));
      i += 2;
      continue;
    }
    bytes.push(input.charCodeAt(i));
  }

  return Buffer.from(bytes);
}

function decodeTransferToBuffer(input, encoding) {
  if (encoding.includes("base64")) {
    return decodeBase64ToBuffer(input);
  }
  if (encoding.includes("quoted-printable")) {
    return decodeQuotedPrintableToBuffer(input);
  }
  return Buffer.from(input, "utf8");
}

function decodeQuotedPrintableToBuffer(input) {
  const normalized = input.replace(/=\r?\n/g, "");
  const bytes = [];

  for (let i = 0; i < normalized.length; i += 1) {
    if (normalized[i] === "=" && /[A-Fa-f0-9]{2}/.test(normalized.slice(i + 1, i + 3))) {
      bytes.push(parseInt(normalized.slice(i + 1, i + 3), 16));
      i += 2;
      continue;
    }
    bytes.push(normalized.charCodeAt(i));
  }

  return Buffer.from(bytes);
}

function decodeBase64ToBuffer(input) {
  const cleaned = input.replace(/[^A-Za-z0-9+/=]/g, "");
  try {
    return Buffer.from(cleaned, "base64");
  } catch {
    return Buffer.from(input, "utf8");
  }
}

function decodeBuffer(buffer, charset) {
  const resolved = (charset || "utf-8").toLowerCase();

  try {
    if (resolved === "utf-8" || resolved === "utf8" || resolved === "us-ascii") {
      return buffer.toString("utf8");
    }
    if (resolved === "iso-8859-1" || resolved === "latin1" || resolved === "windows-1252") {
      return buffer.toString("latin1");
    }

    const decoder = new TextDecoder(resolved, { fatal: false });
    return decoder.decode(buffer);
  } catch {
    return buffer.toString("utf8");
  }
}

function decodeMimeWords(value) {
  return String(value || "").replace(/=\?([^?]+)\?([bBqQ])\?([^?]+)\?=/g, (_, charset, enc, text) => {
    const encoding = enc.toLowerCase();
    const resolvedCharset = (charset || "utf-8").toLowerCase();

    let buffer;
    if (encoding === "b") {
      buffer = decodeBase64ToBuffer(text);
    } else {
      buffer = decodeQuotedPrintableToBuffer(text.replace(/_/g, " "));
    }

    return decodeBuffer(buffer, resolvedCharset);
  });
}

function normalizeContentId(contentId) {
  return String(contentId || "").trim().replace(/^<|>$/g, "").toLowerCase();
}

function defaultAttachmentName(contentType, state) {
  state.attachmentCounter += 1;
  const extension = extensionFromMime(contentType);
  return `attachment-${state.attachmentCounter}${extension}`;
}

function extensionFromMime(contentType) {
  const map = {
    "application/pdf": ".pdf",
    "application/zip": ".zip",
    "application/json": ".json",
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/gif": ".gif",
    "text/plain": ".txt",
    "text/html": ".html"
  };

  return map[contentType] || ".bin";
}

function stripHtmlTags(input) {
  return String(input || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"');
}

function plainToHtml(input) {
  return `<pre>${escapeHtml(input || "")}</pre>`;
}

function compactWhitespace(input) {
  return String(input || "").replace(/\s+/g, " ").trim();
}

function escapeHtml(input) {
  return String(input || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function hash(input) {
  let h = 0;
  for (let i = 0; i < input.length; i += 1) {
    h = (h << 5) - h + input.charCodeAt(i);
    h |= 0;
  }
  return Math.abs(h).toString(36);
}

module.exports = {
  parseMbox,
  parseMessageChunk
};
