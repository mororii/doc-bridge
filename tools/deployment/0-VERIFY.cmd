@echo off
setlocal
chcp 65001 >nul
title DocBridge Package Verification
cd /d "%~dp0"
echo.
echo ============================================================
echo  DocBridge package integrity check
echo ============================================================
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0support\Test-PackageIntegrity.ps1"
set "DOCBRIDGE_EXIT=%ERRORLEVEL%"
echo.
if "%DOCBRIDGE_EXIT%"=="0" (
  echo PASS - Package files are intact.
) else (
  echo FAIL - Do not install this package. Download or copy it again.
)
echo.
if not "%DOCBRIDGE_NO_PAUSE%"=="1" pause
exit /b %DOCBRIDGE_EXIT%
