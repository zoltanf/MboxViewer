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
- creates a GitHub Release
- uploads the built installers/archives to that release

## Manual GitHub UI Flow

If you prefer not to create a tag locally, you can also:

1. Open the GitHub repository.
2. Go to `Actions`.
3. Open the `Release` workflow.
4. Click `Run workflow`.
5. Optionally enter a base version like `1.6`.

This path also builds the artifacts on GitHub and publishes the release automatically.

## Notes

- The helper script checks for tracked working tree changes and refuses to tag if tracked files are dirty.
- Untracked files are ignored by the script.
- The version used by all CI runners is synchronized, so macOS, Windows, and Linux builds all share the same release version.
- If you later add signing or notarization, that can be layered into the workflow with GitHub secrets.
