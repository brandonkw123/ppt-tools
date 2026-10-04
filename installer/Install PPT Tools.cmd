@echo off
rem ============================================================
rem Install PPT Tools
rem Copies PPTTools.ppam (from this folder) into your PowerPoint
rem add-ins folder and registers it so the PPT Tools tab appears
rem when PowerPoint starts. Only changes your own user account;
rem no admin rights needed. Run "Uninstall PPT Tools.cmd" to undo.
rem ============================================================
setlocal
set "SRC=%~dp0PPTTools.ppam"
set "DEST=%APPDATA%\Microsoft\AddIns"
set "KEY=HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTTools"

if not exist "%SRC%" (
    echo PPTTools.ppam was not found next to this installer.
    echo Unzip the whole download first, then run the installer from the unzipped folder.
    goto :done
)

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 (
    echo Please close PowerPoint, then run this installer again.
    goto :done
)

if not exist "%DEST%" mkdir "%DEST%"
copy /Y "%SRC%" "%DEST%\PPTTools.ppam" >nul
if errorlevel 1 (
    echo Could not copy PPTTools.ppam into "%DEST%".
    goto :done
)

rem Clear the "downloaded from the internet" mark so Office doesn't block the macros
powershell -NoProfile -Command "Unblock-File -LiteralPath (Join-Path $env:APPDATA 'Microsoft\AddIns\PPTTools.ppam')"

rem Register the add-in so PowerPoint loads it at startup
reg add "%KEY%" /v Path /t REG_SZ /d "PPTTools.ppam" /f >nul
reg add "%KEY%" /v AutoLoad /t REG_DWORD /d 1 /f >nul

echo PPT Tools is installed. Open PowerPoint and look for the PPT Tools tab.

:done
echo.
pause
