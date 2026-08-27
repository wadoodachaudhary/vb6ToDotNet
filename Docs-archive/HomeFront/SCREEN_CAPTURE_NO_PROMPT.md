# Disable the screen-capture prompt for the Feedback tool (Chrome / Edge)

> ⚠️ **IMPORTANT CORRECTION (verified 2026-06-24): these policies do NOT remove the prompt.**
> Testing showed that with `TabCaptureAllowedByOrigins` / `SameOriginTabCaptureAllowedByOrigins` /
> `ScreenCaptureAllowedByOrigins` all applied (`chrome://policy` → Mandatory / OK), Chrome **still**
> shows "Allow … to see this tab?" on every capture. Those policies only *allow* capture (override an
> admin block); they do **not** bypass the per-capture consent dialog, which is a hard browser security
> boundary with **no supported off switch**. The steps below are kept for reference only — they will not
> make capture prompt-free. Practical guidance: keep the prompt (it's one click, and there's no longer a
> lingering "sharing" bar), or accept html2canvas's limitations for a no-prompt capture.


The Feedback capture (**F8** = full screen, **Shift+F8** = crop) uses the browser's
native screen capture. The browser shows **"Allow … to see this tab?"** the first
time per session — by design, a web page **cannot** suppress that prompt itself.

To remove it on a tester's machine, pre-authorize the app's **origin** with a
browser policy. The app already calls `getDisplayMedia({ preferCurrentTab: true })`,
so a matching policy makes the browser auto-grant the current tab with **no picker
and no prompt**.

> Without this policy you still only get **one** prompt per session (the app reuses
> the share), so the policy is optional — it just removes that last prompt for a
> fully silent experience on managed/tester machines.

---

## Option A — Quick: import the .reg file (Windows, per machine)

1. Open **`EnableScreenCapture-Chrome-Edge.reg`** (in this folder) in Notepad.
2. Replace `http://localhost:5065` with the **exact origin** from the browser
   address bar — scheme + host + port, **no path, no trailing slash**
   (e.g. `http://localhost:5065`, or the production URL `https://homefront.yourserver`).
   Keep only the Chrome or Edge block you use; add more origins as `"2"`, `"3"`, …
3. **Double-click** the file → **Yes** to merge (needs admin for the `HKEY_LOCAL_MACHINE` keys).
4. **Fully quit** the browser (all windows) and reopen it.
5. **Verify:** open `chrome://policy` (or `edge://policy`), find
   **`TabCaptureAllowedByOrigins`**, confirm it lists your origin with status **OK**
   (click *Reload policies* if needed).
6. Open the app and press **F8** — it should capture with no prompt.

What the .reg sets (Chrome shown; Edge is the same under `Microsoft\Edge`):

```
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\TabCaptureAllowedByOrigins]
"1"="http://localhost:5065"
```

**No admin rights?** Use `HKEY_CURRENT_USER` instead of `HKEY_LOCAL_MACHINE` in
both keys — Chrome and Edge also read per-user policies.

---

## Option B — At scale: Group Policy (for IT / domain fleets)

1. Load the **Chrome** or **Microsoft Edge** ADMX administrative templates into
   Group Policy.
2. Go to **Computer (or User) Configuration → Administrative Templates →
   Google Chrome / Microsoft Edge → Content Settings → "Allow Tab capture by these
   sites"** (maps to the `TabCaptureAllowedByOrigins` policy).
3. **Enable** it and add the origin(s), e.g. `http://localhost:5065` or
   `https://homefront.yourserver`.
4. `gpupdate /force`, restart the browser, and verify on `chrome://policy`.

---

## Notes

- **Policy choice.** `TabCaptureAllowedByOrigins` (tab capture only) matches how the
  app captures its own tab and is the least-privilege fit. The tighter
  `SameOriginTabCaptureAllowedByOrigins` also works; the broader
  `ScreenCaptureAllowedByOrigins` is a fallback if a particular browser build still
  prompts.
- **Origin must match exactly** — scheme + host + port. If testers run on different
  ports, add each as its own numbered value.
- **Sharing indicator.** The *prompt* goes away; the browser may still show a small
  "sharing this tab" indicator while a capture is active. That's separate from the
  prompt and is not removable by policy.
- **Secure context.** Native capture needs HTTPS or `localhost` (the app uses both).
