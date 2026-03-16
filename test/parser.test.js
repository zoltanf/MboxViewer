const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { parseMbox, parseMessageChunk } = require("../src/mboxParser");

const FIXTURES_DIR = path.join(__dirname, "fixtures");

test("parses UTF-8 EML fixtures", () => {
  const raw = fs.readFileSync(path.join(FIXTURES_DIR, "sample-utf8.eml"));
  const parsed = parseMessageChunk(raw, {
    index: 1,
    includeAttachmentData: false,
    includeBodyHtml: true
  });

  assert.equal(parsed.subject, "UTF-8 greeting");
  assert.match(parsed.bodyText, /Hello from the UTF-8 fixture/);
  assert.equal(parsed.attachments.length, 0);
});

test("preserves quoted-printable latin1 text until charset decoding", () => {
  const raw = fs.readFileSync(path.join(FIXTURES_DIR, "sample-latin1.eml"));
  const parsed = parseMessageChunk(raw, {
    index: 1,
    includeAttachmentData: false,
    includeBodyHtml: true
  });

  assert.equal(parsed.subject, "Café update");
  assert.match(parsed.bodyText, /café est prêt/i);
});

test("parses multipart messages with inline CID data and attachments", () => {
  const raw = fs.readFileSync(path.join(FIXTURES_DIR, "sample-mailbox.mbox"));
  const messages = parseMbox(raw, {
    includeAttachmentData: "all",
    includeBodyHtml: true
  });
  const htmlMessage = messages.find((message) => message.subject === "HTML with inline image");

  assert.ok(htmlMessage);
  assert.equal(htmlMessage.attachments.length, 2);
  assert.equal(htmlMessage.attachments[0].fileName, "hero.png");
  assert.equal(htmlMessage.attachments[0].isInline, true);
  assert.ok(htmlMessage.attachments[0].base64.length > 0);
  assert.equal(htmlMessage.attachments[1].fileName, "notes.txt");
});

test("tolerates malformed multipart messages without boundaries", () => {
  const raw = Buffer.from(
    [
      "From: Broken <broken@example.test>",
      "To: Reader <reader@example.test>",
      "Subject: Broken multipart",
      "MIME-Version: 1.0",
      'Content-Type: multipart/mixed; charset="utf-8"',
      "",
      "This still has a body even though the MIME boundary is missing."
    ].join("\n"),
    "utf8"
  );

  const parsed = parseMessageChunk(raw, {
    index: 1,
    includeAttachmentData: false,
    includeBodyHtml: true
  });

  assert.equal(parsed.subject, "Broken multipart");
  assert.match(parsed.bodyText, /missing/i);
  assert.equal(parsed.attachments.length, 0);
});
