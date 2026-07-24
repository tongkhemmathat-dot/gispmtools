@echo off
setlocal
cd /d "%~dp0"

rem ---------------------------------------------------------------------
rem  PMtools build launcher
rem  Double-click this file to build the single-file distributable into
rem  Dist\. No Administrator rights needed - the build only reads the
rem  source tree here and writes into Dist\.
rem
rem      Build-PM.cmd                  double-click; builds the .cmd form
rem      Build-PM.cmd -Format Both     also build the .ps1 form
rem      Build-PM.cmd -OutDir D:\Out   write elsewhere instead of Dist\
rem ---------------------------------------------------------------------

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-PMSingle.ps1" %*
set PMEXIT=%errorlevel%

rem Hold the window open only for a double-click. A caller that passed
rem arguments wants the exit code back, not a keypress.
if not "%~1"=="" goto :done

echo.
echo Press any key to close this window.
pause >nul

:done
endlocal & exit /b %PMEXIT%
