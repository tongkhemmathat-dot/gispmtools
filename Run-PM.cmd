@echo off
setlocal
cd /d "%~dp0"

rem ---------------------------------------------------------------------
rem  PMtools launcher
rem  Double-click this file to run the assessment. It sets an execution
rem  policy for its own PowerShell process only (nothing on the machine is
rem  changed) and requests Administrator rights, which several checks need
rem  in order to read machine-wide state.
rem ---------------------------------------------------------------------

net session >nul 2>&1
if %errorlevel%==0 goto :elevated

echo Requesting Administrator privileges...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
exit /b

:elevated
echo.
rem Show-PMMenu.ps1 asks how long to sample first, then runs the assessment.
rem To skip the menu entirely, call Start-PMCheck.ps1 directly instead.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Show-PMMenu.ps1"
set PMEXIT=%errorlevel%

echo.
if %PMEXIT%==0 echo Result: no action items.
if %PMEXIT%==1 echo Result: items requiring monitoring were found - see the report.
if %PMEXIT%==2 echo Result: critical items were found - see the report.
if %PMEXIT% GTR 2 echo Result: the tool did not finish. Review the messages above.

echo.
echo Press any key to close this window.
pause >nul
endlocal
