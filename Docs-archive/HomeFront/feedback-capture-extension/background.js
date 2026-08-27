// HomeFront Feedback Capture — MV3 service worker.
//
// Captures the VISIBLE CONTENT of the requesting tab via chrome.tabs.captureVisibleTab and
// returns it as a PNG data URL. captureVisibleTab needs no screen-share prompt and only ever
// captures the tab's own viewport (never the desktop / other windows). It works with the
// host_permissions for the tab's origin (no `tabs` permission, no user gesture required).
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (!msg || msg.type !== "HF_CAPTURE_TAB") return;

    const windowId = sender.tab && sender.tab.windowId;
    try {
        chrome.tabs.captureVisibleTab(windowId, { format: "png" }, (dataUrl) => {
            const err = chrome.runtime.lastError;
            if (err || !dataUrl) {
                sendResponse({ error: (err && err.message) || "captureVisibleTab returned no data" });
            } else {
                sendResponse({ dataUrl: dataUrl });
            }
        });
    } catch (e) {
        sendResponse({ error: String(e) });
    }
    return true; // keep the message channel open for the async sendResponse
});
