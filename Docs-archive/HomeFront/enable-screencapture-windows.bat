@echo off
REM ============================================================================
REM  HomeFront / Precision Builder - disable the Chrome/Edge screen-capture
REM  "Allow ... to see this tab?" prompt for the Feedback tool (F8 / Shift+F8).
REM
REM  Windows equivalent of the Mac config profile. It pre-authorizes the app's
REM  ORIGIN(s) to capture their OWN browser tab, so getDisplayMedia runs with NO
REM  picker and NO prompt. Writes per-user policies (HKCU) - NO admin needed.
REM
REM  SECURITY NOTE: this only pre-authorizes the SPECIFIC origins listed below
REM  to capture their own tab. It is NOT a blanket "any site can capture the
REM  screen" - other sites still get the normal prompt. Use it only on the
REM  internal testing machines that run the Feedback tool.
REM
REM  HOW TO USE:
REM   1. Edit the ORIGINS list below so it matches the EXACT origin(s) shown in
REM      the browser address bar (scheme + host + port; no path, no trailing
REM      slash), space-separated. Add or remove entries as needed.
REM   2. Double-click this file (or run it from a command prompt).
REM   3. FULLY quit Chrome / Edge (every window) and reopen.
REM   4. Verify at  chrome://policy  (or  edge://policy ) -> "Reload policies":
REM      TabCaptureAllowedByOrigins should list your origins.
REM   5. Open the app and press F8 - it should capture with no prompt.
REM ============================================================================

setlocal EnableDelayedExpansion

REM ---- Origins to pre-authorize (space-separated; edit to match your setup) --
REM   The ORIGIN is scheme + host + port ONLY - drop the /main (or any) path.
REM   The app is served over HTTPS on an explicit port, so keep the port number.
REM   https://98.81.50.253:80  = deployments on port 80  (e.g. https://98.81.50.253:80/main)
REM   https://98.81.50.253:81  = deployments on port 81
REM   http://localhost:5065    = local dev (edit/remove to match your dev URL)
set "ORIGINS=https://98.81.50.253:80 https://98.81.50.253:81 http://localhost:5065"

REM ---- Policy roots: Chrome + Edge, per-user (no admin) ----------------------
set "CHROME=HKCU\Software\Policies\Google\Chrome"
set "EDGE=HKCU\Software\Policies\Microsoft\Edge"

for %%B in ("%CHROME%" "%EDGE%") do (
  for %%P in (TabCaptureAllowedByOrigins SameOriginTabCaptureAllowedByOrigins ScreenCaptureAllowedByOrigins) do (
    set /a N=0
    for %%O in (%ORIGINS%) do (
      set /a N+=1
      reg add "%%~B\%%P" /v !N! /t REG_SZ /d "%%O" /f >nul
    )
  )
)

echo.
echo   Pre-authorized tab capture (Chrome + Edge, per-user) for:
for %%O in (%ORIGINS%) do echo       %%O
echo.
echo   NEXT: fully quit Chrome/Edge (every window) and reopen, then verify at
echo   chrome://policy  ("Reload policies" -^> TabCaptureAllowedByOrigins).
echo   Then press F8 in the app - it should capture with no prompt.
echo.
pause
endlocal
