@echo off
setlocal
chcp 65001 >nul
title DocBridge Installation
cd /d "%~dp0"
echo.
echo ============================================================
echo  DocBridge installation
echo  Close Codex, Claude, Kimi, Cursor, Excel, HWP, and AutoCAD first.
echo ============================================================
echo.
if not exist "%~dp0payload\codex-marketplace\plugins\doc-bridge\dist\doc-bridge-mcp.exe" goto PACKAGE_INCOMPLETE
if not exist "%~dp0support\DocBridge.Deployment.psm1" goto PACKAGE_INCOMPLETE
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-DocBridge.ps1" %*
set "DOCBRIDGE_EXIT=%ERRORLEVEL%"
if not "%DOCBRIDGE_EXIT%"=="0" goto INSTALL_FAILED
echo.
echo ============================================================
echo  INSTALLATION SUCCESS
echo  Next: restart your AI client, then run 2-TEST.cmd.
echo ============================================================
echo.
if not "%DOCBRIDGE_NO_PAUSE%"=="1" pause
exit /b 0

:PACKAGE_INCOMPLETE
set "DOCBRIDGE_EXIT=3"
echo.
echo ============================================================
echo  INSTALLATION NOT STARTED - PACKAGE FILES ARE INCOMPLETE
echo ============================================================
echo  Do not run this file inside the ZIP preview window.
echo  Right-click the ZIP, choose Extract All, then run 1-INSTALL.cmd
echo  from the fully extracted folder.
goto INSTALL_END

:INSTALL_FAILED
echo.
echo ============================================================
echo  INSTALLATION FAILED
echo ============================================================
echo  Do NOT run 2-TEST.cmd yet.
echo  Keep or photograph the error shown above for troubleshooting.
echo  Read the Troubleshooting section in the HTML guide.

:INSTALL_END
echo.
if not "%DOCBRIDGE_NO_PAUSE%"=="1" pause
exit /b %DOCBRIDGE_EXIT%
