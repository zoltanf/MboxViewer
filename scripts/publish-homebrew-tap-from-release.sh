#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if [[ $# -lt 1 || $# -gt 2 ]]; then
  cat >&2 <<'EOF'
Usage: ./scripts/publish-homebrew-tap-from-release.sh <tag|version> [arm64|x64]
Example: ./scripts/publish-homebrew-tap-from-release.sh v1.6.2604160534
Example: ./scripts/publish-homebrew-tap-from-release.sh 1.6.2604160534 x64
EOF
  exit 1
fi

TAG_INPUT="$1"
TARGET_ARCH="${2:-arm64}"

case "$TARGET_ARCH" in
  arm64|x64) ;;
  *)
    echo "Invalid architecture '$TARGET_ARCH'. Use 'arm64' or 'x64'." >&2
    exit 1
    ;;
esac

VERSION="${TAG_INPUT#v}"
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]{10}$ ]]; then
  echo "Invalid tag/version '$TAG_INPUT'. Expected v<major>.<minor>.<YYMMDDHHmm>." >&2
  exit 1
fi
TAG="v$VERSION"

if ! gh auth status >/dev/null 2>&1; then
  echo "gh is not authenticated. Run: gh auth login -h github.com" >&2
  exit 1
fi

ORIGIN_URL="$(git remote get-url origin)"
if [[ "$ORIGIN_URL" =~ github\.com[:/]([^/]+)/([^.]+)(\.git)?$ ]]; then
  SOURCE_REPO="${BASH_REMATCH[1]}/${BASH_REMATCH[2]}"
else
  echo "Could not infer GitHub source repo from origin: $ORIGIN_URL" >&2
  exit 1
fi

echo "Fetching release asset metadata for $SOURCE_REPO@$TAG ($TARGET_ARCH)"
ASSET_LINE="$(
  gh release view "$TAG" --repo "$SOURCE_REPO" --json assets | \
    node -e '
      const arch = process.argv[1];
      const version = process.argv[2];
      const payload = JSON.parse(require("fs").readFileSync(0, "utf8"));
      const assets = Array.isArray(payload.assets) ? payload.assets : [];
      const matcher = arch === "arm64"
        ? (name) => name.includes(`${version}-arm64.dmg`)
        : (name) => name.endsWith(`${version}.dmg`) && !name.includes("arm64");
      const asset = assets.find((entry) => matcher(String(entry.name || "")));
      if (!asset) {
        console.error(`No matching DMG asset found for arch=${arch} version=${version}.`);
        process.exit(2);
      }
      const digest = String(asset.digest || "");
      if (!digest.startsWith("sha256:")) {
        console.error(`Release asset ${asset.name} is missing a sha256 digest.`);
        process.exit(3);
      }
      const sha = digest.slice("sha256:".length);
      process.stdout.write(`${asset.name}\t${sha}`);
    ' "$TARGET_ARCH" "$VERSION"
)"

ASSET_NAME="${ASSET_LINE%%$'\t'*}"
ASSET_SHA="${ASSET_LINE#*$'\t'}"

if [[ -z "$ASSET_NAME" || -z "$ASSET_SHA" || "$ASSET_NAME" == "$ASSET_SHA" ]]; then
  echo "Could not parse release asset metadata for $TAG." >&2
  exit 1
fi

ENV_FILE="$(mktemp "/tmp/mbox-viewer-homebrew-${VERSION}-${TARGET_ARCH}.XXXXXX.env")"
cleanup() {
  rm -f "$ENV_FILE"
}
trap cleanup EXIT

cat > "$ENV_FILE" <<EOF
MBOX_VIEWER_TARGET_ARCH='$TARGET_ARCH'
MBOX_VIEWER_BUILD_VERSION='$VERSION'
MBOX_VIEWER_HOMEBREW_CASK_PATH='/tmp/$ASSET_NAME'
MBOX_VIEWER_HOMEBREW_CASK_SHA256='$ASSET_SHA'
EOF

echo "Publishing Homebrew tap update from release asset: $ASSET_NAME"
"$ROOT/scripts/publish-homebrew-tap.sh" "$ENV_FILE"

echo "Homebrew tap publish completed for $TAG ($TARGET_ARCH)."
