@echo off
setlocal EnableExtensions
REM ===========================================================================
REM  HomeFront Feedback Capture - Windows install helper
REM
REM  Double-click this file. It copies the extension to a stable folder and opens
REM  Chrome's Extensions page + that folder, so you only have to click "Load
REM  unpacked" once. Keep this .bat inside the feedback-capture-extension folder.
REM ===========================================================================

set "SRC=%~dp0"
set "DEST=%LOCALAPPDATA%\HomeFront\feedback-capture-extension"

if not exist "%SRC%manifest.json" (
  echo.
  echo  ERROR: manifest.json was not found next to this script.
  echo  Keep install-windows.bat inside the "feedback-capture-extension" folder and re-run.
  echo.
  pause
  exit /b 1
)

echo.
echo  Installing HomeFront Feedback Capture...
echo    from: %SRC%
echo    to:   %DEST%
echo.

REM Copy the extension files to a stable per-user location (exclude this .bat itself).
robocopy "%SRC%." "%DEST%" /MIR /XF "install-windows.bat" /NFL /NDL /NJH /NJS /NC /NS >nul

REM Open Chrome's Extensions page and the destination folder in Explorer.
start "" chrome "chrome://extensions"
if errorlevel 1 start "" "chrome://extensions"
start "" explorer "%DEST%"

echo  ===========================================================
echo   FINISH IN CHROME  (one time, takes 10 seconds):
echo.
echo     1. Turn ON  "Developer mode"   (toggle, top-right).
echo     2. Click    "Load unpacked".
echo     3. Select this folder:
echo          %DEST%
echo     4. Reload the HomeFront app tab  (Ctrl+F5).
echo.
echo   Then press F8 in the app - it captures with NO prompt.
echo  ===========================================================
echo.
pause
endlocal
