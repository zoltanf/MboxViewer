#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const { parseArgs } = require("util");

const ENV_LINE_RE = /^([A-Za-z_][A-Za-z0-9_]*)=(.*)$/;
const ARCH_ORDER = { arm64: 0, x64: 1 };

function parseShellValue(rawValue) {
  const value = String(rawValue || "").trim();
  if (!value) {
    return "";
  }
  if (value.startsWith("'") && value.endsWith("'")) {
    return value.slice(1, -1).replace(/'"'"'/g, "'");
  }
  if (value.startsWith("\"") && value.endsWith("\"")) {
    return value.slice(1, -1);
  }
  return value;
}

function loadEnvFile(filePath) {
  const entries = {};
  const content = fs.readFileSync(filePath, "utf8");
  for (const rawLine of content.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }
    const match = ENV_LINE_RE.exec(line);
    if (!match) {
      continue;
    }
    entries[match[1]] = parseShellValue(match[2]);
  }
  return entries;
}

function buildArtifact(filePath) {
  const env = loadEnvFile(filePath);
  const requiredKeys = [
    "MBOX_VIEWER_TARGET_ARCH",
    "MBOX_VIEWER_BUILD_VERSION",
    "MBOX_VIEWER_HOMEBREW_CASK_PATH",
    "MBOX_VIEWER_HOMEBREW_CASK_SHA256",
  ];
  for (const key of requiredKeys) {
    if (!env[key]) {
      throw new Error(`${filePath} is missing expected build metadata: ${key}`);
    }
  }
  return {
    arch: env.MBOX_VIEWER_TARGET_ARCH,
    version: env.MBOX_VIEWER_BUILD_VERSION,
    caskFilename: path.basename(env.MBOX_VIEWER_HOMEBREW_CASK_PATH),
    caskSha256: env.MBOX_VIEWER_HOMEBREW_CASK_SHA256,
  };
}

function renderCask({
  artifacts,
  sourceRepo,
  caskToken = "mbox-viewer",
  appName = "Mbox Viewer",
  desc = "Desktop viewer for .mbox, .eml, and .pst email archives",
  homepage = "https://github.com/zoltanf/MboxViewer",
}) {
  if (!artifacts.length) {
    throw new Error("No artifacts were provided.");
  }

  const version = artifacts[0].version;
  const lines = [
    `cask "${caskToken}" do`,
    `  version "${version}"`,
  ];

  if (artifacts.length === 2) {
    const byArch = Object.fromEntries(artifacts.map((artifact) => [artifact.arch, artifact]));
    lines.push('  arch arm: "arm64", intel: "x64"');
    lines.push(`  sha256 arm: "${byArch.arm64.caskSha256}", intel: "${byArch.x64.caskSha256}"`);
    lines.push(`  url "https://github.com/${sourceRepo}/releases/download/v#{version}/${caskToken}-homebrew-#{version}-#{arch}.tar.gz"`);
  } else {
    const artifact = artifacts[0];
    const archDependency = artifact.arch === "x64" ? "x86_64" : artifact.arch;
    lines.push(`  sha256 "${artifact.caskSha256}"`);
    lines.push(`  url "https://github.com/${sourceRepo}/releases/download/v#{version}/${artifact.caskFilename}"`);
    lines.push(`  depends_on arch: :${archDependency}`);
  }

  lines.push(
    `  name "${appName}"`,
    `  desc "${desc}"`,
    `  homepage "${homepage}"`,
    "",
    `  app "${appName}.app"`,
    "",
    "  caveats do",
    "    <<~EOS",
    "      If macOS blocks the first launch because the build is not notarized yet, remove the quarantine flag:",
    `        sudo xattr -r -d com.apple.quarantine "/Applications/${appName}.app"`,
    "    EOS",
    "  end",
    "end",
    "",
  );

  return lines.join("\n");
}

function renderReadme({
  sourceRepo,
  tapRepo,
  caskToken = "mbox-viewer",
}) {
  return [
    "# Mbox Viewer Homebrew Tap",
    "",
    "Install the app with:",
    "",
    "```bash",
    `brew tap ${tapRepo}`,
    `brew install --cask ${caskToken}`,
    "```",
    "",
    "If macOS blocks the first launch because the build is not notarized yet, remove the quarantine flag:",
    "",
    "```bash",
    'sudo xattr -r -d com.apple.quarantine "/Applications/Mbox Viewer.app"',
    "```",
    "",
    `Artifacts are published from https://github.com/${sourceRepo}.`,
    "",
  ].join("\n");
}

function writeTap({
  tapDir,
  sourceRepo,
  tapRepo,
  artifactsEnv,
  caskToken = "mbox-viewer",
  appName = "Mbox Viewer",
  desc = "Desktop viewer for .mbox, .eml, and .pst email archives",
  homepage = "https://github.com/zoltanf/MboxViewer",
}) {
  const artifacts = artifactsEnv
    .map((filePath) => buildArtifact(path.resolve(filePath)))
    .sort((left, right) => (ARCH_ORDER[left.arch] ?? 99) - (ARCH_ORDER[right.arch] ?? 99));

  if (!artifacts.length) {
    throw new Error("No build artifacts were provided.");
  }

  const versions = new Set(artifacts.map((artifact) => artifact.version));
  if (versions.size !== 1) {
    throw new Error(`All artifacts must have the same version, got: ${Array.from(versions).join(", ")}`);
  }

  const arches = artifacts.map((artifact) => artifact.arch);
  if (new Set(arches).size !== arches.length) {
    throw new Error(`Duplicate architectures provided: ${arches.join(", ")}`);
  }

  const absoluteTapDir = path.resolve(tapDir);
  const casksDir = path.join(absoluteTapDir, "Casks");
  fs.mkdirSync(casksDir, { recursive: true });

  fs.writeFileSync(
    path.join(casksDir, `${caskToken}.rb`),
    renderCask({ artifacts, sourceRepo, caskToken, appName, desc, homepage }),
    "utf8"
  );

  fs.writeFileSync(
    path.join(absoluteTapDir, "README.md"),
    renderReadme({ sourceRepo, tapRepo, caskToken }),
    "utf8"
  );
}

function main(argv = process.argv.slice(2)) {
  const { values } = parseArgs({
    args: argv,
    options: {
      "tap-dir": { type: "string" },
      "source-repo": { type: "string" },
      "tap-repo": { type: "string" },
      "artifacts-env": { type: "string", multiple: true },
      "cask-token": { type: "string", default: "mbox-viewer" },
      "app-name": { type: "string", default: "Mbox Viewer" },
      desc: { type: "string", default: "Desktop viewer for .mbox, .eml, and .pst email archives" },
      homepage: { type: "string", default: "https://github.com/zoltanf/MboxViewer" },
    },
    allowPositionals: false,
  });

  for (const required of ["tap-dir", "source-repo", "tap-repo"]) {
    if (!values[required]) {
      throw new Error(`Missing required option --${required}`);
    }
  }

  if (!values["artifacts-env"] || values["artifacts-env"].length === 0) {
    throw new Error("Provide at least one --artifacts-env value.");
  }

  writeTap({
    tapDir: values["tap-dir"],
    sourceRepo: values["source-repo"],
    tapRepo: values["tap-repo"],
    artifactsEnv: values["artifacts-env"],
    caskToken: values["cask-token"],
    appName: values["app-name"],
    desc: values.desc,
    homepage: values.homepage,
  });
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(error.message || String(error));
    process.exit(1);
  }
}

module.exports = {
  buildArtifact,
  loadEnvFile,
  parseShellValue,
  renderCask,
  renderReadme,
  writeTap,
};
