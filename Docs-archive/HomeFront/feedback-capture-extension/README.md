# HomeFront Feedback Capture — Chrome/Edge extension

A tiny, optional browser extension that lets the Feedback tool (F8 / Shift+F8) capture the
**current tab** with **no screen-share prompt and no "sharing this tab" bar**. It uses
`chrome.tabs.captureVisibleTab`, which captures only the tab's own visible content — never the
desktop or other windows.

**Optional by design:** when the extension is installed the app uses it automatically (silent
capture). When it isn't, the app falls back to the normal `getDisplayMedia` path (the one-click
prompt). No app change is needed to switch between them.

---

## Install for a tester (Load unpacked)

1. Copy the **`feedback-capture-extension`** folder to the tester's machine.
2. Open **`chrome://extensions`** (or `edge://extensions`).
3. Turn on **Developer mode** (top-right).
4. Click **Load unpacked** and select the `feedback-capture-extension` folder.
5. It appears as **"HomeFront Feedback Capture"**. Done.
6. Open the app and press **F8** — the capture happens with no prompt.

> Keep the folder's files together (`manifest.json`, `background.js`, `content.js`, plus the
> helper `install-windows.bat` and this `README.md`).

### Windows testers — run the helper

On Windows, just **double-click `install-windows.bat`** (inside this folder). It copies the
extension to `%LOCALAPPDATA%\HomeFront\feedback-capture-extension` and opens Chrome's Extensions
page and that folder for you — then do the one-time **Developer mode → Load unpacked → select the
folder** step it prints, and reload the app. (Chrome requires that final click for any unpacked
extension; it can't be scripted. For a fully hands-off fleet install, use `ExtensionInstallForcelist`
— see below.)

## Which sites it runs on

`host_permissions` is `<all_urls>` — `chrome.tabs.captureVisibleTab` **requires** it; a scoped host
permission (e.g. `http://localhost/*`) is **not** sufficient for that API and the capture silently
fails. The extension only *injects* its content script on the origins in `content_scripts.matches`,
which default to **`http://localhost`** / **`https://localhost`** (any port). To use it on a deployed
server, add that origin to **`matches`** and reload the extension:

```jsonc
"content_scripts": [{
  "matches": [
    "http://localhost/*",
    "https://localhost/*",
    "https://homefront.yourserver.com/*"      // ← add your production origin here
  ],
  ...
}]
```

## Roll out to many machines (optional)

For a fleet, host the packed extension and force-install it with the Chrome/Edge
**`ExtensionInstallForcelist`** policy (this enterprise policy *does* take effect, unlike the
capture-allow policies). That installs it silently for every user and prevents them disabling it.
Ask if you want the packaging + policy steps.

## How it works (for maintainers)

```
feedback.js (page)  --postMessage-->  content.js  --chrome.runtime-->  background.js
        ^                                                                    |
        |               PNG data URL of the visible tab                      |
        +--------------------- postMessage <----------- captureVisibleTab ---+
```

- `feedback.js` pings on init; if the extension answers, it routes F8/Shift+F8 capture through it.
- `background.js` calls `chrome.tabs.captureVisibleTab({format:'png'})` and returns the data URL.
- Uses the `<all_urls>` host permission (required by `captureVisibleTab`) and no user gesture; the
  content script is scoped to the `matches` origins, so the extension only ever captures the app's tab.

## Privacy / security

- Captures **only** when the app page asks (F8/Shift+F8), and **only** the visible tab content.
- The content script runs only on the `matches` origins, so it only acts on the app's own tab.
- `host_permissions` is `<all_urls>` purely because `captureVisibleTab` mandates it; the extension
  still captures only when the app page asks (F8/Shift+F8), and only the visible tab.
