// HomeFront Feedback Capture — content-script bridge.
//
// The app page (feedback.js) cannot call extension APIs directly, so it talks to this script via
// window.postMessage, and this script relays to the background service worker via chrome.runtime.
// Only same-window messages of our own known types are handled.
(function () {
    "use strict";

    function announce() {
        // Tell the page the no-prompt capture extension is available.
        window.postMessage({ type: "HF_FEEDBACK_CAPTURE_AVAILABLE" }, "*");
    }

    window.addEventListener("message", function (ev) {
        if (ev.source !== window || !ev.data) return;
        var d = ev.data;

        if (d.type === "HF_FEEDBACK_CAPTURE_PING") {
            announce();
            return;
        }

        if (d.type === "HF_FEEDBACK_CAPTURE_REQUEST") {
            var reqId = d.reqId;
            try {
                chrome.runtime.sendMessage({ type: "HF_CAPTURE_TAB" }, function (resp) {
                    var err = chrome.runtime.lastError ? chrome.runtime.lastError.message
                                                       : (resp && resp.error);
                    window.postMessage({
                        type: "HF_FEEDBACK_CAPTURE_RESPONSE",
                        reqId: reqId,
                        dataUrl: (resp && resp.dataUrl) || null,
                        error: err || ((resp && resp.dataUrl) ? null : "no data")
                    }, "*");
                });
            } catch (e) {
                window.postMessage({
                    type: "HF_FEEDBACK_CAPTURE_RESPONSE",
                    reqId: reqId, dataUrl: null, error: String(e)
                }, "*");
            }
        }
    });

    // Announce on load (the page may also PING us once feedback.js initialises).
    announce();
})();
