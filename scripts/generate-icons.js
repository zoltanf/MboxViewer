const { execFileSync, spawnSync } = require("node:child_process");
const { mkdirSync, copyFileSync, rmSync, existsSync } = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const buildDir = path.join(root, "build");
const svgPath = path.join(root, "assets", "app-icon.svg");
const basePng = path.join(buildDir, "icon-1024.png");
const iconPng = path.join(buildDir, "icon.png");
const iconIco = path.join(buildDir, "icon.ico");

const sizes = [16, 24, 32, 48, 64, 128, 256, 512, 1024];

function hasCommand(name) {
  const result = spawnSync("which", [name], { stdio: "ignore" });
  return result.status === 0;
}

function run(cmd, args) {
  execFileSync(cmd, args, { stdio: "inherit" });
}

function resizePng(input, output, size) {
  if (hasCommand("sips")) {
    run("sips", ["-z", String(size), String(size), input, "--out", output]);
    return;
  }

  if (hasCommand("magick")) {
    run("magick", [input, "-resize", `${size}x${size}`, output]);
    return;
  }

  throw new Error("Need either 'sips' or 'magick' to resize icons.");
}

function generateBasePng() {
  mkdirSync(buildDir, { recursive: true });

  if (hasCommand("rsvg-convert")) {
    run("rsvg-convert", ["-w", "1024", "-h", "1024", "-o", basePng, svgPath]);
    return;
  }

  if (hasCommand("magick")) {
    run("magick", ["-background", "none", svgPath, "-resize", "1024x1024", basePng]);
    return;
  }

  throw new Error("Need either 'rsvg-convert' or 'magick' to rasterize the SVG icon.");
}

function generatePngSet() {
  copyFileSync(basePng, iconPng);
  for (const size of sizes) {
    const filePath = path.join(buildDir, `icon-${size}.png`);
    if (size === 1024) {
      copyFileSync(basePng, filePath);
    } else {
      resizePng(basePng, filePath, size);
    }
  }
}

function generateIco() {
  if (!hasCommand("magick")) {
    console.warn("Skipping ICO generation: 'magick' not found.");
    return;
  }

  const inputs = [16, 24, 32, 48, 64, 128, 256].map((size) => path.join(buildDir, `icon-${size}.png`));
  run("magick", [...inputs, iconIco]);
}

function assertOutputs() {
  const required = [iconPng, iconIco];
  const missing = required.filter((filePath) => !existsSync(filePath));
  if (missing.length > 0) {
    throw new Error(`Icon generation incomplete. Missing: ${missing.join(", ")}`);
  }
}

function main() {
  rmSync(path.join(buildDir, "icon.icns"), { force: true });
  rmSync(path.join(buildDir, "icon.iconset"), { recursive: true, force: true });
  generateBasePng();
  generatePngSet();
  generateIco();
  assertOutputs();
  console.log("Generated app icons in build/ (icon.png, icon.ico).");
}

main();
