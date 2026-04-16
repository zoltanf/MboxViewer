#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "This build script only runs on macOS." >&2
  exit 1
fi

APP_NAME="${MBOX_VIEWER_APP_NAME:-Mbox Viewer}"
CASK_TOKEN="${MBOX_VIEWER_CASK_TOKEN:-mbox-viewer}"

if [[ -n "${MBOX_VIEWER_TARGET_ARCHES:-}" ]]; then
  TARGET_ARCH_INPUT="$MBOX_VIEWER_TARGET_ARCHES"
elif [[ -n "${MBOX_VIEWER_TARGET_ARCH:-}" ]]; then
  TARGET_ARCH_INPUT="$MBOX_VIEWER_TARGET_ARCH"
elif [[ "$(uname -m)" == "arm64" ]]; then
  TARGET_ARCH_INPUT="arm64 x64"
else
  TARGET_ARCH_INPUT="x64"
fi
TARGET_ARCH_INPUT="${TARGET_ARCH_INPUT//,/ }"

TARGET_ARCHES=()
for arch in $TARGET_ARCH_INPUT; do
  case "$arch" in
    arm64|x64) TARGET_ARCHES+=("$arch") ;;
    *)
      echo "Unsupported target architecture: $arch" >&2
      exit 1
      ;;
  esac
done

if [[ ${#TARGET_ARCHES[@]} -eq 0 ]]; then
  echo "Set MBOX_VIEWER_TARGET_ARCHES or MBOX_VIEWER_TARGET_ARCH to at least one architecture." >&2
  exit 1
fi

if [[ -n "${MBOX_VIEWER_BUILD_VERSION:-}" ]]; then
  if [[ ! "$MBOX_VIEWER_BUILD_VERSION" =~ ^([0-9]+\.[0-9]+)\.([0-9]{10})$ ]]; then
    echo "MBOX_VIEWER_BUILD_VERSION must look like <major>.<minor>.<YYMMDDHHmm>." >&2
    exit 1
  fi
  MBOX_VIEWER_BASE_VERSION="${BASH_REMATCH[1]}"
  export MBOX_VIEWER_BUILD_STAMP="${BASH_REMATCH[2]}"
fi

if [[ -n "${MBOX_VIEWER_BASE_VERSION:-}" ]]; then
  bash "$ROOT/set-base-version.sh" "$MBOX_VIEWER_BASE_VERSION"
fi

DIST_ROOT="$ROOT/dist"
HOMEBREW_DIST_ROOT="$DIST_ROOT/macos"
BUILD_ROOT="$ROOT/build/macos"
SHA256SUMS_PATH="$HOMEBREW_DIST_ROOT/SHA256SUMS.txt"
ARTIFACTS_ENV="$BUILD_ROOT/artifacts.env"
ARTIFACTS_LIST="$BUILD_ROOT/artifacts.list"

rm -rf "$BUILD_ROOT" "$HOMEBREW_DIST_ROOT"
mkdir -p "$BUILD_ROOT" "$HOMEBREW_DIST_ROOT"

BUILD_CMD=(npm run dist:mac --)
for arch in "${TARGET_ARCHES[@]}"; do
  BUILD_CMD+=("--${arch}")
done

echo "Building macOS release artifacts for: ${TARGET_ARCHES[*]}"
"${BUILD_CMD[@]}"

VERSION="$(node -p "require('./package.json').version")"
if [[ ! "$VERSION" =~ ^([0-9]+\.[0-9]+)\.([0-9]{10})$ ]]; then
  echo "package.json version '$VERSION' does not match the expected release format." >&2
  exit 1
fi

BASE_VERSION="${BASH_REMATCH[1]}"
BUILD_STAMP="${BASH_REMATCH[2]}"

quote_shell() {
  printf "%s" "$1" | sed "s/'/'\"'\"'/g"
}

write_env_line() {
  local key="$1"
  local value="$2"
  printf "%s='%s'\n" "$key" "$(quote_shell "$value")"
}

find_app_bundle() {
  local target_arch="$1"
  local -a candidates=()
  if [[ "$target_arch" == "arm64" ]]; then
    candidates=(
      "$DIST_ROOT/mac-arm64/$APP_NAME.app"
      "$DIST_ROOT/mac-arm64-unpacked/$APP_NAME.app"
    )
  else
    candidates=(
      "$DIST_ROOT/mac/$APP_NAME.app"
      "$DIST_ROOT/mac-x64/$APP_NAME.app"
      "$DIST_ROOT/mac-unpacked/$APP_NAME.app"
    )
  fi

  local candidate
  for candidate in "${candidates[@]}"; do
    if [[ -d "$candidate" ]]; then
      printf "%s\n" "$candidate"
      return 0
    fi
  done

  echo "Could not find built app bundle for $target_arch." >&2
  exit 1
}

score_release_file() {
  local target_arch="$1"
  local filename="$2"
  if [[ "$target_arch" == "arm64" ]]; then
    [[ "$filename" == *arm64* ]] && printf "0\n" || printf "99\n"
    return 0
  fi

  if [[ "$filename" == *arm64* ]] || [[ "$filename" == *universal* ]]; then
    printf "99\n"
  elif [[ "$filename" == *x64* ]] || [[ "$filename" == *x86_64* ]]; then
    printf "0\n"
  else
    printf "10\n"
  fi
}

find_release_file() {
  local target_arch="$1"
  local extension="$2"
  local best_path=""
  local best_score=999
  local best_length=999
  local candidate

  while IFS= read -r candidate; do
    local filename
    local score
    local length

    filename="$(basename "$candidate")"
    if [[ "$filename" != *"$VERSION"* ]]; then
      continue
    fi

    score="$(score_release_file "$target_arch" "$filename")"
    if [[ "$score" -ge 99 ]]; then
      continue
    fi

    length="${#filename}"
    if (( score < best_score || (score == best_score && length < best_length) )); then
      best_path="$candidate"
      best_score="$score"
      best_length="$length"
    fi
  done < <(find "$DIST_ROOT" -maxdepth 1 -type f -name "*.${extension}" -print | sort)

  if [[ -z "$best_path" ]]; then
    echo "Could not find a .$extension release artifact for $target_arch and version $VERSION." >&2
    exit 1
  fi

  printf "%s\n" "$best_path"
}

: > "$SHA256SUMS_PATH"
: > "$ARTIFACTS_LIST"

ARCH_ARTIFACTS_ENVS=()
for target_arch in "${TARGET_ARCHES[@]}"; do
  ARCH_BUILD_ROOT="$BUILD_ROOT/$target_arch"
  APP_PATH="$(find_app_bundle "$target_arch")"
  DMG_PATH="$(find_release_file "$target_arch" "dmg")"
  ZIP_PATH="$(find_release_file "$target_arch" "zip")"
  HOMEBREW_STAGE="$ARCH_BUILD_ROOT/homebrew-cask"
  HOMEBREW_CASK_PATH="$HOMEBREW_DIST_ROOT/${CASK_TOKEN}-homebrew-${VERSION}-${target_arch}.tar.gz"
  ARCH_ARTIFACTS_ENV="$ARCH_BUILD_ROOT/artifacts.env"

  mkdir -p "$ARCH_BUILD_ROOT"
  rm -rf "$HOMEBREW_STAGE"
  mkdir -p "$HOMEBREW_STAGE"

  ditto "$APP_PATH" "$HOMEBREW_STAGE/$APP_NAME.app"
  tar -C "$HOMEBREW_STAGE" -czf "$HOMEBREW_CASK_PATH" "$APP_NAME.app"

  DMG_SHA256="$(shasum -a 256 "$DMG_PATH" | awk '{print $1}')"
  ZIP_SHA256="$(shasum -a 256 "$ZIP_PATH" | awk '{print $1}')"
  HOMEBREW_CASK_SHA256="$(shasum -a 256 "$HOMEBREW_CASK_PATH" | awk '{print $1}')"

  cat >> "$SHA256SUMS_PATH" <<EOF
${DMG_SHA256}  $(basename "$DMG_PATH")
${ZIP_SHA256}  $(basename "$ZIP_PATH")
${HOMEBREW_CASK_SHA256}  $(basename "$HOMEBREW_CASK_PATH")
EOF

  {
    write_env_line "MBOX_VIEWER_BUILD_VERSION" "$VERSION"
    write_env_line "MBOX_VIEWER_BASE_VERSION" "$BASE_VERSION"
    write_env_line "MBOX_VIEWER_BUILD_STAMP" "$BUILD_STAMP"
    write_env_line "MBOX_VIEWER_TARGET_ARCH" "$target_arch"
    write_env_line "MBOX_VIEWER_APP_NAME" "$APP_NAME"
    write_env_line "MBOX_VIEWER_CASK_TOKEN" "$CASK_TOKEN"
    write_env_line "MBOX_VIEWER_APP_PATH" "$APP_PATH"
    write_env_line "MBOX_VIEWER_DMG_PATH" "$DMG_PATH"
    write_env_line "MBOX_VIEWER_DMG_SHA256" "$DMG_SHA256"
    write_env_line "MBOX_VIEWER_ZIP_PATH" "$ZIP_PATH"
    write_env_line "MBOX_VIEWER_ZIP_SHA256" "$ZIP_SHA256"
    write_env_line "MBOX_VIEWER_HOMEBREW_CASK_PATH" "$HOMEBREW_CASK_PATH"
    write_env_line "MBOX_VIEWER_HOMEBREW_CASK_SHA256" "$HOMEBREW_CASK_SHA256"
    write_env_line "MBOX_VIEWER_SHA256SUMS_PATH" "$SHA256SUMS_PATH"
  } > "$ARCH_ARTIFACTS_ENV"

  ARCH_ARTIFACTS_ENVS+=("$ARCH_ARTIFACTS_ENV")
  printf "%s\n" "$ARCH_ARTIFACTS_ENV" >> "$ARTIFACTS_LIST"
done

ARTIFACTS_ENV_LIST_VALUE=""
TARGET_ARCHES_VALUE=""
for i in "${!TARGET_ARCHES[@]}"; do
  if [[ -n "$ARTIFACTS_ENV_LIST_VALUE" ]]; then
    ARTIFACTS_ENV_LIST_VALUE="${ARTIFACTS_ENV_LIST_VALUE}:"
    TARGET_ARCHES_VALUE="${TARGET_ARCHES_VALUE} "
  fi
  ARTIFACTS_ENV_LIST_VALUE="${ARTIFACTS_ENV_LIST_VALUE}${ARCH_ARTIFACTS_ENVS[$i]}"
  TARGET_ARCHES_VALUE="${TARGET_ARCHES_VALUE}${TARGET_ARCHES[$i]}"
done

{
  write_env_line "MBOX_VIEWER_BUILD_VERSION" "$VERSION"
  write_env_line "MBOX_VIEWER_BASE_VERSION" "$BASE_VERSION"
  write_env_line "MBOX_VIEWER_BUILD_STAMP" "$BUILD_STAMP"
  write_env_line "MBOX_VIEWER_TARGET_ARCHES" "$TARGET_ARCHES_VALUE"
  write_env_line "MBOX_VIEWER_ARTIFACTS_LIST_PATH" "$ARTIFACTS_LIST"
  write_env_line "MBOX_VIEWER_ARTIFACTS_ENV_LIST" "$ARTIFACTS_ENV_LIST_VALUE"
  write_env_line "MBOX_VIEWER_SHA256SUMS_PATH" "$SHA256SUMS_PATH"
  for i in "${!TARGET_ARCHES[@]}"; do
    write_env_line "MBOX_VIEWER_ARTIFACTS_ENV_${TARGET_ARCHES[$i]}" "${ARCH_ARTIFACTS_ENVS[$i]}"
  done
} > "$ARTIFACTS_ENV"

echo
echo "Build complete."
echo "  version: $VERSION"
for target_arch in "${TARGET_ARCHES[@]}"; do
  echo "  ${target_arch} metadata: $BUILD_ROOT/$target_arch/artifacts.env"
done
echo "  checksums: $SHA256SUMS_PATH"
echo "  manifest: $ARTIFACTS_ENV"
echo "  list: $ARTIFACTS_LIST"
