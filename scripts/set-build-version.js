const { readFileSync, writeFileSync } = require("fs");
const path = require("path");

const workspaceRoot = path.join(__dirname, "..");
const packageJsonPath = path.join(workspaceRoot, "package.json");
const packageLockPath = path.join(workspaceRoot, "package-lock.json");

const packageJson = JSON.parse(readFileSync(packageJsonPath, "utf8"));
const packageLock = JSON.parse(readFileSync(packageLockPath, "utf8"));

const [major = "1", minor = "0"] = String(packageJson.version || "1.0.0").split(".");
const buildStamp = formatBuildStamp(new Date());
const nextVersion = `${major}.${minor}.${buildStamp}`;

packageJson.version = nextVersion;
packageLock.version = nextVersion;
if (packageLock.packages && packageLock.packages[""]) {
  packageLock.packages[""].version = nextVersion;
}

writeJson(packageJsonPath, packageJson);
writeJson(packageLockPath, packageLock);

console.log(`Updated build version to ${nextVersion}`);

function formatBuildStamp(date) {
  const year = String(date.getFullYear()).slice(-2);
  const month = pad2(date.getMonth() + 1);
  const day = pad2(date.getDate());
  const hours = pad2(date.getHours());
  const minutes = pad2(date.getMinutes());
  return `${year}${month}${day}${hours}${minutes}`;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function writeJson(filePath, value) {
  writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}
