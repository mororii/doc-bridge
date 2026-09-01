@echo off
setlocal
chcp 65001 >nul
title DocBridge Uninstall
cd /d "%~dp0"
echo.
echo ============================================================
echo  DocBridge uninstall
echo ============================================================
echo.
if not "%DOCBRIDGE_ASSUME_YES%"=="1" (
  choice /C YN /N /M "Remove DocBridge settings and HWP security registration? [Y/N] "
  if errorlevel 2 exit /b 0
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-DocBridge.ps1" -RemoveHwpSecurity %*
set "DOCBRIDGE_EXIT=%ERRORLEVEL%"
echo.
if "%DOCBRIDGE_EXIT%"=="0" (
  echo Uninstall completed. Backups and reports were retained.
) else (
  echo Uninstall reported an error. See the message above.
)
echo.
if not "%DOCBRIDGE_NO_PAUSE%"=="1" pause
exit /b %DOCBRIDGE_EXIT%
