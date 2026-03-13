# Mbox Viewer

Mbox Viewer is a desktop app for browsing `.mbox`, `.eml`, and Outlook `.pst` email files without importing them into a mail client.

It is built with Electron and designed for fast local exploration of mailbox dumps, backups, and exports.

## What This App Does

- Opens a single `.mbox`, `.eml`, or `.pst` file from your computer
- Builds a local SQLite sidecar index (`<your-file>.mbox.sqlite`) for fast search and navigation
- Shows messages in a two-pane UI:
  - left: sender, subject, date, preview
  - right: full rendered message content
- Supports message attachments and per-message `.eml` export

Everything runs locally on your machine.

## Key Features

- Fast indexing with progress feedback for large mailboxes
- Index reuse on later opens when the source file is unchanged
- Full-text search (subject, sender, recipient, snippet, body, attachment names)
- Date range filtering with a dual-handle slider (in a filter popover)
- Pagination for large result sets
- Inline CID image resolution for HTML emails
- Remote-content blocking toggle for HTML email, disabled by default for privacy
- Downloadable attachments
- Direct viewing of standalone `.eml` message files without creating a SQLite index
- Export selected message as `.eml`
- PST support via local conversion to a cached sidecar mbox (`<file>.pst.mbox`) before indexing, including message attachments

## Security-Oriented Behavior

- Email HTML is sanitized before rendering
- Potentially unsafe schemes are stripped from email content (`javascript:`, etc.)
- Remote images and other remote HTML content can be blocked before rendering
  - the remote-content toolbar toggle is hidden until a mailbox is open
  - remote content is blocked by default
  - when content is blocked, the message view shows a privacy notice
- Clicking a link in an email shows a confirmation dialog before opening externally
  - full URL is shown
  - the registrable domain is highlighted for easier visual verification
- External links open in the system default browser only after confirmation

## Why SQLite Sidecar Indexing?

`.mbox` files can be very large. Building a local index gives:

- much faster search
- faster page loads while browsing
- persistent performance across sessions

The generated index is stored next to your mbox file and reused when valid.

## Getting Started

### Requirements

- Node.js + npm
- macOS / Windows / Linux supported by Electron

### Run in Development

```bash
npm install
npm start
```

Note: `npm install` runs `electron-builder install-app-deps` to rebuild native modules for your Electron version.

### Install on macOS with Homebrew

```bash
brew install --cask zoltanf/mboxviewer/mbox-viewer
```

If macOS blocks the app on first launch because it is not notarized, open:

- `System Settings`
- `Privacy & Security`

Then scroll to the bottom and click `Open Anyway` for `Mbox Viewer.app`.

### Open `.mbox`, `.eml`, and `.pst` files directly from Finder or Explorer

Packaged builds register Mbox Viewer as a viewer for `.mbox`, `.eml`, and `.pst` files, so you can open email files by double-clicking them.

Notes:

- this works with the packaged app, not with `npm start`
- after installing a new build, you may need to choose Mbox Viewer once as the default app for these file types

On macOS:

- move `Mbox Viewer.app` to `/Applications`
- select an `.mbox`, `.eml`, or `.pst` file in Finder
- press `Cmd+I`
- under `Open with`, choose `Mbox Viewer`
- click `Change All...` if you want to use it for all files of that type

On Windows:

- install the packaged app
- right-click an `.mbox`, `.eml`, or `.pst` file
- choose `Open with`
- choose `Mbox Viewer`
- enable `Always use this app` if you want it as the default

## Build / Package

Build outputs go to `dist/`.

### Unpacked app (quick local package)

```bash
npm run pack
```

### Platform-specific installers

```bash
npm run dist:mac
npm run dist:win
npm run dist:linux
```

### All configured targets

```bash
npm run dist
```

### Build Pipeline Note

`pack`/`dist` scripts trigger:

- automatic build-version update using `major.minor.YYMMDDHHmm`
- icon generation from `assets/app-icon.svg`

Example build version: `1.4.2603131526`

If you are packaging frequently, be aware that the version in `package.json` and `package-lock.json` will change on each build.

## Repository Structure (high level)

- `main.js` - Electron main process + IPC
- `preload.js` - secure renderer API bridge
- `src/renderer/` - UI (HTML/CSS/renderer logic)
- `src/mboxParser.js` - MIME/message parsing
- `src/mboxStore.js` - SQLite indexing/search/load layer

## License

This project is licensed under the GNU General Public License v3.0.
See [LICENSE](/Users/zoltanf/Development/MboxViewer/LICENSE).
