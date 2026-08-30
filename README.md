# InboxRedux

**Flexible message previews for Outlook on the web.**

InboxRedux gives you a better way to read mail without leaving your inbox. Choose a resizable bottom preview pane or an experimental inline reader that opens a message beneath the row you double-click.

> InboxRedux is an independent project and is not affiliated with, endorsed by, or sponsored by Microsoft.

## Features

- **Resizable Preview Pane** — Keep Outlook's bottom reading pane at the height you prefer. Drag the divider whenever you want; InboxRedux remembers the size.
- **Accordion Reader** *(experimental)* — Double-click a message to open a compact preview directly below its row.
- **Normal selection stays normal** — Single-click, Ctrl-click, and Shift-click retain Outlook's usual message-selection behavior.
- **Local preferences only** — Your selected mode and preferred pane height are stored in Firefox, on your device.

## Privacy and permissions

InboxRedux runs only on `outlook.office.com` and `outlook.office365.com`.

It does not send, collect, store remotely, or sell email content or personal data. It has no network, identity, clipboard, download, or account permissions. The extension uses Firefox's local extension storage only for your selected preview mode and pane height.

## Reader modes

### Preview Pane

Uses Outlook's bottom reading pane with a saved height. This is the default and most stable mode.

### Accordion Reader

Hides the bottom pane. Double-click a message row to place a compact, read-only copy of the message below that row. Click **×**, or double-click that same row again, to close it.

The Accordion Reader is experimental because Outlook on the web is a frequently updated application. Use Outlook's normal open or pop-out controls for reply, forwarding, attachments, downloads, and other message actions.

## Install for development or testing

1. Download or clone this repository.
2. In Firefox, open `about:debugging#/runtime/this-firefox`.
3. Choose **Load Temporary Add-on**.
4. Select this repository's `manifest.json`.
5. Open Outlook on the web and open InboxRedux's Preferences from `about:addons`.

Temporary add-ons are removed when Firefox restarts. A signed release from Firefox Add-ons will be available after review.

## Development

This is a no-build Firefox Manifest V3 extension:

- `manifest.json` — extension configuration and permissions
- `content.js` — Outlook layout and reader-mode behavior
- `accordion.css` — inline-reader presentation
- `options.html` / `options.js` — mode and pane-height preferences

## Feedback

Please open a GitHub issue with your Firefox version, the Outlook on the web layout you use, and a screenshot if InboxRedux does not behave as expected.
