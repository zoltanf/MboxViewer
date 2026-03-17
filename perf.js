const os = require("os");
const path = require("path");
const { mkdtemp, cp, rm } = require("fs/promises");
const {
  ensureMboxDatabase,
  ensurePstDatabase,
  searchMessages,
  loadMessageById
} = require("../src/mboxStore");

const FIXTURE_DIR = process.env.MBOX_VIEWER_PERF_FIXTURES || path.join(__dirname, "..", "test", "fixtures");

function createMockSender() {
  return {
    isDestroyed: () => false,
    send: () => {}
  };
}

async function withWorkspace(fileName, runner) {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), "mbox-viewer-perf-"));
  const sourcePath = path.join(tempRoot, fileName);
  const dbPath = `${sourcePath}.sqlite`;

  try {
    await cp(path.join(FIXTURE_DIR, fileName), sourcePath);
    return await runner({ sourcePath, dbPath });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

async function measureFixture(kind, fileName) {
  return withWorkspace(fileName, async ({ sourcePath, dbPath }) => {
    const sender = createMockSender();
    const startedAt = performance.now();
    const indexed =
      kind === "pst"
        ? await ensurePstDatabase(sourcePath, sender, { dbPath })
        : await ensureMboxDatabase(sourcePath, sender, { dbPath });
    const indexReadyMs = Math.round(performance.now() - startedAt);

    const firstPageStartedAt = performance.now();
    const firstPage = searchMessages(dbPath, "", 20, 0);
    const firstPageReadyMs = Math.round(performance.now() - firstPageStartedAt);

    const firstMessageStartedAt = performance.now();
    const firstMessage = firstPage.messages[0] ? await loadMessageById(dbPath, firstPage.messages[0].id) : null;
    const firstMessageReadyMs = Math.round(performance.now() - firstMessageStartedAt);

    const reuseStartedAt = performance.now();
    const reused =
      kind === "pst"
        ? await ensurePstDatabase(sourcePath, sender, { dbPath })
        : await ensureMboxDatabase(sourcePath, sender, { dbPath });
    const reuseOpenMs = Math.round(performance.now() - reuseStartedAt);

    return {
      fixture: fileName,
      kind,
      indexed,
      firstPageCount: firstPage.total,
      firstMessageSubject: firstMessage?.subject || null,
      timingsMs: {
        indexReady: indexReadyMs,
        firstPageReady: firstPageReadyMs,
        firstMessageReady: firstMessageReadyMs,
        reuseOpen: reuseOpenMs
      },
      reuseResult: reused
    };
  });
}

async function main() {
  const results = [];
  results.push(await measureFixture("mbox", "sample-mailbox.mbox"));
  results.push(await measureFixture("mbox", "sample-medium.mbox"));
  results.push(await measureFixture("pst", "sample-enron.pst"));
  console.log(JSON.stringify({ generatedAt: new Date().toISOString(), results }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
