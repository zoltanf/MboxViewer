#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "Usage: $0 <major.minor>" >&2
  echo "Example: $0 1.5" >&2
  exit 1
fi

BASE_VERSION="$1"

if [[ ! "$BASE_VERSION" =~ ^[0-9]+\.[0-9]+$ ]]; then
  echo "Invalid version '$BASE_VERSION'. Expected format: major.minor" >&2
  exit 1
fi

FULL_VERSION="${BASE_VERSION}.0"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

node - "$FULL_VERSION" "$REPO_ROOT" <<'EOF'
const { readFileSync, writeFileSync } = require("fs");
const path = require("path");

const nextVersion = process.argv[2];
const workspaceRoot = process.argv[3];
const packageJsonPath = path.join(workspaceRoot, "package.json");
const packageLockPath = path.join(workspaceRoot, "package-lock.json");

const packageJson = JSON.parse(readFileSync(packageJsonPath, "utf8"));
const packageLock = JSON.parse(readFileSync(packageLockPath, "utf8"));

packageJson.version = nextVersion;
packageLock.version = nextVersion;
if (packageLock.packages && packageLock.packages[""]) {
  packageLock.packages[""].version = nextVersion;
}

writeFileSync(packageJsonPath, `${JSON.stringify(packageJson, null, 2)}\n`, "utf8");
writeFileSync(packageLockPath, `${JSON.stringify(packageLock, null, 2)}\n`, "utf8");

console.log(`Updated base version to ${nextVersion}`);
EOF
