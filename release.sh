#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "Usage: ./release.sh <major.minor>" >&2
  echo "Example: ./release.sh 1.6" >&2
  exit 1
fi

BASE_VERSION="${1#v}"

if [[ ! "$BASE_VERSION" =~ ^[0-9]+\.[0-9]+$ ]]; then
  echo "Invalid version '$1'. Expected format: major.minor" >&2
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

if ! git diff --quiet || ! git diff --cached --quiet; then
  echo "Refusing to create a release tag with unstaged or staged tracked changes." >&2
  echo "Commit or stash your tracked changes first." >&2
  exit 1
fi

BUILD_STAMP="$(date -u +'%y%m%d%H%M')"
RELEASE_TAG="v${BASE_VERSION}.${BUILD_STAMP}"

if git rev-parse "$RELEASE_TAG" >/dev/null 2>&1; then
  echo "Tag '$RELEASE_TAG' already exists locally." >&2
  exit 1
fi

if git ls-remote --exit-code --tags origin "refs/tags/${RELEASE_TAG}" >/dev/null 2>&1; then
  echo "Tag '$RELEASE_TAG' already exists on origin." >&2
  exit 1
fi

git tag -a "$RELEASE_TAG" -m "Release $RELEASE_TAG"
git push origin "$RELEASE_TAG"

cat <<EOF
Created and pushed tag: $RELEASE_TAG

GitHub Actions will now:
- run tests
- build macOS / Windows / Linux artifacts
- create a GitHub Release for $RELEASE_TAG
- upload the generated release assets

You can follow progress in:
https://github.com/zoltanf/MboxViewer/actions
EOF
