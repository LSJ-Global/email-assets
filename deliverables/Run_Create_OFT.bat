@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File ".\Create_OFT_Windows_Outlook.ps1"
echo.
echo ----------------------------------------
echo If the file was created successfully,
echo check this folder for:
echo F5_AppWorld_Seoul_2026_EDM.oft
echo ----------------------------------------
echo.
pause
