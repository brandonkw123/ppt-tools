@echo off
rem ============================================================
rem Uninstall PPT Tools
rem Removes PPTTools.ppam from your PowerPoint add-ins folder and
rem unregisters it. Only changes your own user account.
rem ============================================================
setlocal
set "DEST=%APPDATA%\Microsoft\AddIns\PPTTools.ppam"
set "KEY=HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTTools"

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 (
    echo Please close PowerPoint, then run this again.
    goto :done
)

reg delete "%KEY%" /f >nul 2>&1
if exist "%DEST%" del "%DEST%"

echo PPT Tools is uninstalled.

:done
echo.
pause
