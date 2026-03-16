const test = require("node:test");
const assert = require("node:assert/strict");
const { readFile, utimes } = require("fs/promises");
const {
  ensureMboxDatabase,
  ensurePstDatabase,
  searchMessages,
  loadMessageById,
  getAttachmentData,
  getMessageEmlBuffer,
  getMessageSourcePreview
} = require("../src/mboxStore");
const os = require("os");
const path = require("path");
const { mkdtemp, cp, rm } = require("fs/promises");

async function createFixtureWorkspace(fixtureName) {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), "mbox-viewer-test-"));
  const sourcePath = path.join(tempRoot, fixtureName);
  const fixturePath = path.join(__dirname, "fixtures", fixtureName);
  await cp(fixturePath, sourcePath);

  return {
    tempRoot,
    sourcePath,
    dbPath: `${sourcePath}.sqlite`,
    async cleanup() {
      await rm(tempRoot, { recursive: true, force: true });
    }
  };
}

function createMockSender() {
  return {
    isDestroyed: () => false,
    send: () => {}
  };
}

test("mbox indexing, reuse, and on-demand attachment loading work", async () => {
  const workspace = await createFixtureWorkspace("sample-mailbox.mbox");
  try {
    const sender = createMockSender();
    const first = await ensureMboxDatabase(workspace.sourcePath, sender, { dbPath: workspace.dbPath });
    assert.equal(first.totalMessages, 2);
    assert.equal(first.reused, false);

    const reused = await ensureMboxDatabase(workspace.sourcePath, sender, { dbPath: workspace.dbPath });
    assert.equal(reused.reused, true);

    const page = searchMessages(workspace.dbPath, "", 10, 0);
    assert.equal(page.total, 2);
    assert.equal(page.messages[0].subject, "HTML with inline image");

    const message = await loadMessageById(workspace.dbPath, 2);
    assert.equal(message.attachments.length, 2);
    assert.equal(message.attachments[0].fileName, "hero.png");
    assert.ok(message.attachments[0].base64.length > 0);
    assert.equal(message.attachments[1].fileName, "notes.txt");
    assert.equal(message.attachments[1].base64, "");

    const attachment = await getAttachmentData(workspace.dbPath, 2, message.attachments[1].id);
    assert.equal(attachment.fileName, "notes.txt");
    assert.match(attachment.data.toString("utf8"), /attachment fixture/i);

    const emlBuffer = await getMessageEmlBuffer(workspace.dbPath, 2);
    assert.match(emlBuffer.toString("utf8"), /^From: HTML Sender/m);
  } finally {
    await workspace.cleanup();
  }
});

test("mbox reindexes when the source mtime changes", async () => {
  const workspace = await createFixtureWorkspace("sample-medium.mbox");
  try {
    const sender = createMockSender();
    await ensureMboxDatabase(workspace.sourcePath, sender, { dbPath: workspace.dbPath });
    const now = new Date();
    const later = new Date(now.getTime() + 60_000);
    await utimes(workspace.sourcePath, later, later);

    const reindexed = await ensureMboxDatabase(workspace.sourcePath, sender, { dbPath: workspace.dbPath });
    assert.equal(reindexed.reused, false);
  } finally {
    await workspace.cleanup();
  }
});

test("direct PST indexing loads messages and attachments without a temp mbox sidecar", async () => {
  const workspace = await createFixtureWorkspace("sample-enron.pst");
  try {
    const sender = createMockSender();
    const indexed = await ensurePstDatabase(workspace.sourcePath, sender, { dbPath: workspace.dbPath });
    assert.equal(indexed.totalMessages, 71);
    assert.equal(indexed.reused, false);

    const searchPage = searchMessages(workspace.dbPath, "OBA", 5, 0);
    assert.equal(searchPage.total, 2);
    assert.equal(searchPage.messages[0].subject, "New OBA's");
    assert.equal(searchPage.messages[0].hasAttachments, true);

    const message = await loadMessageById(workspace.dbPath, 1);
    assert.equal(message.subject, "New OBA's");
    assert.equal(message.attachments.length, 11);
    assert.equal(message.attachments[0].fileName, "OBA_27602_Citizens.doc");
    assert.equal(message.attachments[0].base64, "");

    const attachment = await getAttachmentData(workspace.dbPath, 1, message.attachments[0].id);
    assert.equal(attachment.fileName, "OBA_27602_Citizens.doc");
    assert.ok(attachment.data.length > 50_000);

    const sourcePreview = await getMessageSourcePreview(workspace.dbPath, 1);
    assert.match(sourcePreview, /Message-ID: <pst-/);
    await assert.rejects(readFile(`${workspace.sourcePath}.mbox`), /ENOENT/);
  } finally {
    await workspace.cleanup();
  }
});
