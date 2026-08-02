@echo off
setlocal EnableExtensions
title XERO Bar - install / update

set "DESTDIR=%USERPROFILE%\XERO Bar"
set "DEST=%DESTDIR%\XERO_bar.ahk"
set "URL=https://raw.githubusercontent.com/Timeless15000/xero-apps/main/XERO_bar.ahk"
set "ICOURL=https://raw.githubusercontent.com/Timeless15000/xero-apps/main/xero.ico"

echo.
echo   Installing / updating the XERO Bar...
echo.

if not exist "%DESTDIR%" mkdir "%DESTDIR%"

rem  Download the latest bar (cache-busted)
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -UseBasicParsing -Uri ('%URL%?v=' + (Get-Random)) -OutFile '%DEST%'; exit 0 } catch { exit 1 }"
if errorlevel 1 (
  echo.
  echo   Download FAILED. Check your internet connection and run this again.
  echo.
  pause
  exit /b 1
)

rem  Xero logo for the desktop shortcut (best effort)
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri ('%ICOURL%?v=' + (Get-Random)) -OutFile '%DESTDIR%\xero.ico' } catch {}"

rem  Point the desktop shortcut at the new self-updating bar (with Xero icon when available)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $l=$w.CreateShortcut([Environment]::GetFolderPath('Desktop')+'\XERO Bar.lnk'); $l.TargetPath='%DEST%'; $l.WorkingDirectory='%DESTDIR%'; if(Test-Path '%DESTDIR%\xero.ico'){$l.IconLocation='%DESTDIR%\xero.ico,0'}; $l.Save()"
ie4uinit.exe -show >nul 2>nul

echo   Done. Starting the bar...
start "" "%DEST%"

timeout /t 3 >nul
exit /b 0
