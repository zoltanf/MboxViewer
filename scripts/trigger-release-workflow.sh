#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if [[ $# -lt 1 || $# -gt 3 ]]; then
  cat >&2 <<'EOF'
Usage: ./scripts/trigger-release-workflow.sh <major.minor> [draft|publish] [stable|prerelease]
Example: ./scripts/trigger-release-workflow.sh 1.6
Example: ./scripts/trigger-release-workflow.sh 1.6 publish prerelease
EOF
  exit 1
fi

BASE_VERSION="${1#v}"
if [[ ! "$BASE_VERSION" =~ ^[0-9]+\.[0-9]+$ ]]; then
  echo "Invalid version '$1'. Expected format: major.minor" >&2
  exit 1
fi

RELEASE_MODE="${2:-draft}"
CHANNEL_MODE="${3:-stable}"

case "$RELEASE_MODE" in
  draft) DRAFT_VALUE="true" ;;
  publish) DRAFT_VALUE="false" ;;
  *)
    echo "Invalid release mode '$RELEASE_MODE'. Use 'draft' or 'publish'." >&2
    exit 1
    ;;
esac

case "$CHANNEL_MODE" in
  stable) PRERELEASE_VALUE="false" ;;
  prerelease) PRERELEASE_VALUE="true" ;;
  *)
    echo "Invalid channel mode '$CHANNEL_MODE'. Use 'stable' or 'prerelease'." >&2
    exit 1
    ;;
esac

if ! gh auth status >/dev/null 2>&1; then
  echo "gh is not authenticated. Run: gh auth login -h github.com" >&2
  exit 1
fi

echo "Triggering Release workflow on main"
echo "  base_version: $BASE_VERSION"
echo "  draft: $DRAFT_VALUE"
echo "  prerelease: $PRERELEASE_VALUE"

gh workflow run release.yml \
  --ref main \
  -f base_version="$BASE_VERSION" \
  -f draft="$DRAFT_VALUE" \
  -f prerelease="$PRERELEASE_VALUE"

echo
echo "Waiting for the workflow run to appear..."
sleep 3

RUN_INFO="$(gh run list --workflow release.yml --limit 1 --json databaseId,url,status,conclusion)"
RUN_ID="$(echo "$RUN_INFO" | node -e 'const runs = JSON.parse(require("fs").readFileSync(0, "utf8")); if (!runs.length) process.exit(1); process.stdout.write(String(runs[0].databaseId || ""));')"
RUN_URL="$(echo "$RUN_INFO" | node -e 'const runs = JSON.parse(require("fs").readFileSync(0, "utf8")); if (!runs.length) process.exit(1); process.stdout.write(String(runs[0].url || ""));')"

if [[ -z "$RUN_ID" ]]; then
  echo "Could not determine workflow run id. Check manually with: gh run list --workflow release.yml" >&2
  exit 1
fi

if [[ -n "$RUN_URL" ]]; then
  echo "Workflow URL: $RUN_URL"
fi

echo "Streaming logs until completion..."
gh run watch "$RUN_ID" --exit-status
