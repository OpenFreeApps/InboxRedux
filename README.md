# InboxRedux

**Flexible message previews for Outlook on the web, available for Firefox and Chrome.**

InboxRedux gives you a better way to read mail without leaving your inbox. Choose a resizable bottom preview pane or an experimental inline reader that opens a message beneath the row you double-click.

> InboxRedux is an independent project and is not affiliated with, endorsed by, or sponsored by Microsoft.

## Features

- **Resizable Preview Pane** — Keep Outlook's bottom reading pane at the height you prefer. Drag the divider whenever you want; InboxRedux remembers the size.
- **Accordion Reader** *(experimental)* — Double-click a message to open a compact preview directly below its row.
- **Normal selection stays normal** — Single-click, Ctrl-click, and Shift-click retain Outlook's usual message-selection behavior.
- **Firefox and Chrome** — One shared codebase produces separate, store-ready browser packages.
- **Local preferences only** — Your selected mode and preferred pane height are stored in your browser, on your device.

## Privacy and permissions

InboxRedux runs only on `outlook.office.com` and `outlook.office365.com`.

It does not send, collect, store remotely, or sell email content or personal data. It has no network, identity, clipboard, download, or account permissions. The extension uses local extension storage only for your selected preview mode and pane height.

## Reader modes

### Preview Pane

Uses Outlook's bottom reading pane with a saved height. This is the default and most stable mode.

### Accordion Reader

Hides the bottom pane. Double-click a message row to place a compact, read-only copy of the message below that row. Click **×**, or double-click that same row again, to close it.

The Accordion Reader is experimental because Outlook on the web is a frequently updated application. Use Outlook's normal open or pop-out controls for reply, forwarding, attachments, downloads, and other message actions.

## Install for development or testing

### Firefox

1. Download or clone this repository.
2. Run the Firefox build below.
3. In Firefox, open `about:debugging#/runtime/this-firefox`.
4. Choose **Load Temporary Add-on**.
5. Select `builds/firefox/manifest.json`.
6. Open Outlook on the web and open InboxRedux's Preferences from `about:addons`.

Temporary add-ons are removed when Firefox restarts. A signed release from Firefox Add-ons will be available after review.

### Chrome

1. Download or clone this repository.
2. Run the Chrome build below.
3. Open `chrome://extensions`.
4. Enable **Developer mode**.
5. Choose **Load unpacked** and select the generated `builds/chrome` folder.
6. Open Outlook on the web and open InboxRedux's extension details to access its options.

Unpacked Chrome extensions remain installed until you remove them, but Chrome will show a developer-mode warning.

## Build packages

The source code is shared. The build script creates browser-specific manifests so Chrome never receives Firefox-only Gecko metadata.

Requirements: Node.js 18+ and the `zip` command.

```bash
npm run build
```

This creates:

- `builds/firefox/` and `builds/InboxRedux-firefox.xpi` — for Firefox Add-ons submission or testing
- `builds/chrome/` and `builds/InboxRedux-chrome.zip` — for Chrome Web Store submission or testing

To build one target only:

```bash
npm run build:firefox
npm run build:chrome
```

The generated `builds/` directory is intentionally excluded from Git.

## Project structure

```text
manifest.json            Shared Manifest V3 definition
manifests/
  firefox.json           Firefox-only Gecko metadata
content.js               Outlook layout and reader-mode behavior
accordion.css            Inline-reader presentation
options.html
options.js
scripts/
  build.mjs              Produces Firefox and Chrome manifests/packages
builds/                  Generated, ignored build output
```

## Feedback

Please open a GitHub issue with your browser and version, the Outlook on the web layout you use, and a screenshot if InboxRedux does not behave as expected.
