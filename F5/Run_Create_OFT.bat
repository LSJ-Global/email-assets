@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

echo Creating F5 AppWorld Seoul 2026 Outlook template...
echo Folder: %SCRIPT_DIR%
echo.

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Create_OFT_Windows_Outlook.ps1"

echo.
if exist "%SCRIPT_DIR%F5_AppWorld_Seoul_2026_EDM.oft" (
  echo Done: "%SCRIPT_DIR%F5_AppWorld_Seoul_2026_EDM.oft"
) else (
  echo Failed. Please copy the error message above.
)

echo.
pause
