#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if [[ $# -ne 1 ]]; then
  echo "Usage: ./scripts/build-macos-and-publish.sh <major.minor>" >&2
  echo "Example: ./scripts/build-macos-and-publish.sh 1.6" >&2
  exit 1
fi

BASE_VERSION="${1#v}"
if [[ ! "$BASE_VERSION" =~ ^[0-9]+\.[0-9]+$ ]]; then
  echo "Invalid version '$1'. Expected format: major.minor" >&2
  exit 1
fi

export MBOX_VIEWER_BASE_VERSION="$BASE_VERSION"

echo "Starting full macOS + Homebrew release flow for base version ${BASE_VERSION}"
echo "Step 1/3: Build macOS artifacts and Homebrew package tarballs"
"$ROOT/scripts/build-macos.sh"

echo "Step 2/3: Publish release assets to GitHub release"
"$ROOT/scripts/publish-github-release.sh"

echo "Step 3/3: Publish cask update to Homebrew tap"
"$ROOT/scripts/publish-homebrew-tap.sh"

echo "Done. Release flow completed."
