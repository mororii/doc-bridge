@echo off
setlocal
chcp 65001 >nul
title DocBridge Test
cd /d "%~dp0"
echo.
echo ============================================================
echo  DocBridge installation test
echo ============================================================
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-DocBridge.ps1" %*
set "DOCBRIDGE_EXIT=%ERRORLEVEL%"
echo.
if "%DOCBRIDGE_EXIT%"=="0" (
  echo PASS - DocBridge is ready.
) else (
  echo FAIL - Check the lines above and doctor-report.json.
)
echo.
if not "%DOCBRIDGE_NO_PAUSE%"=="1" pause
exit /b %DOCBRIDGE_EXIT%
