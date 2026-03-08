# Kumon Extensions (Chrome)

Chrome extension build of Kumon Extensions: Auto Grader + Worksheet Setter (4-3-3, 3-2, 2-2) for Class-Navi.

## Load in Chrome

1. Open `chrome://extensions/`
2. Enable **Developer mode** (top right)
3. Click **Load unpacked**
4. Select the `extension` folder (this directory)

## Files

- `manifest.json` – Manifest V3, content script + host permissions, storage
- `content.js` – UI, Auto Grader, worksheet dropdown, payload builder (runs in isolated world)
- `inject.js` – Fetch/XHR intercept (injected into page context; reads `body.dataset` and listens for `KumonPayloadReady`)

## Version

Matches userscript v0.5.1.
