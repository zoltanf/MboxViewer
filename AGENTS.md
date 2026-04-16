# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Commands

```bash
npm start                 # Run app in development (Electron)
npm test                  # Run all tests (Node built-in test runner)
npm run test:perf         # Run performance benchmarks
npm run pack              # Quick local package (unpacked, for smoke-testing)
npm run dist:mac          # Build macOS DMG + ZIP
npm run dist:win          # Build Windows NSIS + ZIP
npm run dist:linux        # Build Linux AppImage + DEB
npm run icons:generate    # Regenerate icons from assets/app-icon.svg
```

Run a single test file:
```bash
node --test test/parser.test.js
node --test test/store.test.js
```

There is no lint command configured.

## Architecture

MboxViewer is an Electron desktop app with the standard three-layer structure:

**Main process** (`main.js`) — owns all file I/O, SQLite access, and native OS integration. Exposes functionality to the renderer exclusively via IPC handlers (e.g. `open-mailbox-path`, `search-messages`, `get-message`, `save-attachment`). Never accessed directly from the renderer.

**Preload** (`preload.js`) — thin context-bridge that exposes `window.mboxApi` to the renderer. All renderer↔main communication goes through this surface. Keep it minimal; add new capabilities here only when a new IPC channel is added in `main.js`.

**Renderer** (`src/renderer/`) — vanilla HTML/CSS/JS, no framework. `renderer.js` handles the full UI: two-pane layout (message list + detail), toolbar, filter popover, search with 220 ms debounce, pagination (PAGE_SIZE=200), modals, and bookmark toggling.

**Core modules** (all consumed by main process only):
- `src/mboxParser.js` — MIME parsing, charset decoding, attachment extraction for `.mbox` and `.eml` files
- `src/mboxStore.js` — SQLite sidecar index creation, full-text search, bookmarks, and caching. The index file lives at `<source-file>.sqlite` next to the opened mailbox
- `src/pstConverter.js` — PST file parsing via `pst-extractor`, converts to the same index records used by `mboxStore`

**Security invariants to preserve:**
- `contextIsolation: true`, `nodeIntegration: false` — never relax these
- All HTML from email content must be sanitized before rendering
- External links require a confirmation dialog (`open-external` IPC channel)
- Remote content (images, etc.) is blocked by default; the toggle is in the renderer toolbar

## Testing

Tests use Node's built-in `node:test` module. Fixtures live in `test/fixtures/`. Store tests use `createFixtureWorkspace()` to spin up a temp SQLite workspace; parser tests work directly on fixture files.

## Release

`./release.sh <major.minor>` tags and triggers the GitHub Actions workflow (`.github/workflows/release.yml`). See `RELEASING.md` for the full workflow. Version format: `major.minor.YYMMDDHHmm`.
