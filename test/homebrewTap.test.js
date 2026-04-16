const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const {
  buildArtifact,
  renderCask,
  renderReadme,
  writeTap,
} = require("../scripts/render_homebrew_tap.js");

test("buildArtifact reads per-arch Homebrew metadata", () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "mbox-viewer-homebrew-"));
  const envPath = path.join(tempDir, "artifacts.env");
  fs.writeFileSync(
    envPath,
    [
      "MBOX_VIEWER_TARGET_ARCH='arm64'",
      "MBOX_VIEWER_BUILD_VERSION='1.6.2603172145'",
      "MBOX_VIEWER_HOMEBREW_CASK_PATH='/tmp/mbox-viewer-homebrew-1.6.2603172145-arm64.tar.gz'",
      "MBOX_VIEWER_HOMEBREW_CASK_SHA256='abc123'",
      "",
    ].join("\n"),
    "utf8"
  );

  assert.deepEqual(buildArtifact(envPath), {
    arch: "arm64",
    version: "1.6.2603172145",
    caskFilename: "mbox-viewer-homebrew-1.6.2603172145-arm64.tar.gz",
    caskSha256: "abc123",
  });
});

test("buildArtifact restores shell-escaped apostrophes", () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "mbox-viewer-homebrew-"));
  const envPath = path.join(tempDir, "artifacts.env");
  fs.writeFileSync(
    envPath,
    [
      "MBOX_VIEWER_TARGET_ARCH='x64'",
      "MBOX_VIEWER_BUILD_VERSION='1.6.2603172145'",
      "MBOX_VIEWER_HOMEBREW_CASK_PATH='/tmp/Zoltan'\"'\"'s/mbox-viewer-homebrew-1.6.2603172145-x64.tar.gz'",
      "MBOX_VIEWER_HOMEBREW_CASK_SHA256='def456'",
      "",
    ].join("\n"),
    "utf8"
  );

  assert.deepEqual(buildArtifact(envPath), {
    arch: "x64",
    version: "1.6.2603172145",
    caskFilename: "mbox-viewer-homebrew-1.6.2603172145-x64.tar.gz",
    caskSha256: "def456",
  });
});

test("renderCask emits dual-arch Homebrew URLs", () => {
  const output = renderCask({
    sourceRepo: "zoltanf/MboxViewer",
    artifacts: [
      {
        arch: "arm64",
        version: "1.6.2603172145",
        caskFilename: "mbox-viewer-homebrew-1.6.2603172145-arm64.tar.gz",
        caskSha256: "arm-sha",
      },
      {
        arch: "x64",
        version: "1.6.2603172145",
        caskFilename: "mbox-viewer-homebrew-1.6.2603172145-x64.tar.gz",
        caskSha256: "intel-sha",
      },
    ],
  });

  assert.match(output, /arch arm: "arm64", intel: "x64"/);
  assert.match(output, /sha256 arm: "arm-sha", intel: "intel-sha"/);
  assert.match(output, /mbox-viewer-homebrew-#\{version\}-#\{arch\}\.tar\.gz/);
});

test("writeTap renders the cask and README", () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "mbox-viewer-tap-"));
  const envDir = fs.mkdtempSync(path.join(os.tmpdir(), "mbox-viewer-env-"));
  const envPath = path.join(envDir, "artifacts.env");

  fs.writeFileSync(
    envPath,
    [
      "MBOX_VIEWER_TARGET_ARCH='x64'",
      "MBOX_VIEWER_BUILD_VERSION='1.6.2603172145'",
      "MBOX_VIEWER_HOMEBREW_CASK_PATH='/tmp/mbox-viewer-homebrew-1.6.2603172145-x64.tar.gz'",
      "MBOX_VIEWER_HOMEBREW_CASK_SHA256='intel-sha'",
      "",
    ].join("\n"),
    "utf8"
  );

  writeTap({
    tapDir: tempDir,
    sourceRepo: "zoltanf/MboxViewer",
    tapRepo: "zoltanf/homebrew-mboxviewer",
    artifactsEnv: [envPath],
  });

  const caskPath = path.join(tempDir, "Casks", "mbox-viewer.rb");
  const readmePath = path.join(tempDir, "README.md");

  assert.equal(fs.existsSync(caskPath), true);
  assert.equal(fs.existsSync(readmePath), true);
  assert.match(fs.readFileSync(caskPath, "utf8"), /cask "mbox-viewer" do/);
  assert.match(fs.readFileSync(readmePath, "utf8"), /brew tap zoltanf\/homebrew-mboxviewer/);
});

test("renderReadme points to the source repository", () => {
  const output = renderReadme({
    sourceRepo: "zoltanf/MboxViewer",
    tapRepo: "zoltanf/homebrew-mboxviewer",
  });

  assert.match(output, /Artifacts are published from https:\/\/github\.com\/zoltanf\/MboxViewer\./);
});
