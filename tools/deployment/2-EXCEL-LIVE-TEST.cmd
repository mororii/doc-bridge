@echo off
setlocal
chcp 65001 >nul
title DocBridge Excel Live RCW Test
cd /d "%~dp0"
echo.
echo ============================================================
echo  DocBridge Excel live RCW test
echo  Open an Excel workbook before continuing.
echo ============================================================
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-DocBridge.ps1" -RequireExcelRuntime %*
set "DOCBRIDGE_EXIT=%ERRORLEVEL%"
echo.
if "%DOCBRIDGE_EXIT%"=="0" (
  echo PASS - Excel repeated context is stable.
) else (
  echo FAIL - Open an Excel workbook and check the lines above.
)
echo.
if not "%DOCBRIDGE_NO_PAUSE%"=="1" pause
exit /b %DOCBRIDGE_EXIT%
