# Releasing

The project now supports automated GitHub releases through GitHub Actions.

You do not need to build locally before publishing a release. The GitHub workflow runs tests, builds the platform artifacts, creates the GitHub Release, and uploads the generated files.

## Recommended Release Flow

Use the helper script from the repository root:

```bash
./release.sh 1.6
```

What it does:

- validates the base version (`major.minor`)
- generates a UTC build stamp in the existing `YYMMDDHHmm` format
- creates a tag like `v1.6.2603172145`
- pushes that tag to `origin`
- triggers the GitHub `Release` workflow automatically

## What Happens On GitHub

The workflow in [.github/workflows/release.yml](/Users/zoltanf/Development/MboxViewer/.github/workflows/release.yml):

- validates and normalizes the release version
- runs the test suite
- builds macOS, Windows, and Linux artifacts
- builds `x64` and `arm64` binaries for Windows and Linux on GitHub-hosted runners
- generates Homebrew cask tarballs plus `SHA256SUMS.txt` from the macOS build
- creates a GitHub Release
- uploads the built installers/archives to that release

You can also trigger the workflow directly from your terminal:

```bash
./scripts/trigger-release-workflow.sh 1.6
```

Optional modes:

```bash
./scripts/trigger-release-workflow.sh 1.6 draft stable
./scripts/trigger-release-workflow.sh 1.6 publish prerelease
```

## Manual GitHub UI Flow

If you prefer not to create a tag locally, you can also:

1. Open the GitHub repository.
2. Go to `Actions`.
3. Open the `Release` workflow.
4. Click `Run workflow`.
5. Optionally enter a base version like `1.6`.

This path also builds the artifacts on GitHub and publishes the release automatically.

## Post-release Homebrew Tap Update

After the GitHub Release is published, update the Homebrew tap with:

```bash
./scripts/publish-homebrew-tap-from-release.sh v1.6.2604160534
```

This script:

- reads the DMG asset name and `sha256` from the GitHub release
- generates temporary metadata
- calls `publish-homebrew-tap.sh` for you

Optional architecture override:

```bash
./scripts/publish-homebrew-tap-from-release.sh v1.6.2604160534 x64
```

## Notes

- The helper script checks for tracked working tree changes and refuses to tag if tracked files are dirty.
- Untracked files are ignored by the script.
- The version used by all CI runners is synchronized, so macOS, Windows, and Linux builds all share the same release version.
- The macOS release assets now include Homebrew cask tarballs, so the tap can point at GitHub Release assets directly.
- If you later add signing or notarization, that can be layered into the workflow with GitHub secrets.

## Local Homebrew Maintenance Flow

If you want the same Homebrew publishing workflow used in `NanoHarness`, run these commands on macOS from the repository root:

```bash
./scripts/build-macos.sh
./scripts/publish-github-release.sh
./scripts/publish-homebrew-tap.sh
```

This flow:

- builds the macOS app for the configured architectures
- creates `mbox-viewer-homebrew-<version>-<arch>.tar.gz`
- writes metadata to `build/macos/artifacts.env` and `build/macos/artifacts.list`
- uploads the macOS release assets to `v<version>`
- renders and pushes the Homebrew cask to the tap repo

By default the tap repo is inferred as `<github-owner>/homebrew-mboxviewer`. Override it with:

```bash
export MBOX_VIEWER_TAP_REPO="your-user/homebrew-mboxviewer"
```
