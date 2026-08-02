@echo off
setlocal EnableExtensions
title OUTLOOK Bar - install / update

set "URL=https://raw.githubusercontent.com/Timeless15000/xero-apps/main/OUTLOOK_bar.ahk"

rem  Find the Outlook folder (the one that contains src\index.js)
set "DESTDIR="
for %%D in (
  "%OneDrive%\Document\GitHub\Outlook"
  "%OneDrive%\Documents\GitHub\Outlook"
  "%USERPROFILE%\OneDrive\Document\GitHub\Outlook"
  "%USERPROFILE%\OneDrive\Documents\GitHub\Outlook"
  "%USERPROFILE%\Documents\GitHub\Outlook"
  "%USERPROFILE%\Document\GitHub\Outlook"
  "%USERPROFILE%\GitHub\Outlook"
  "%USERPROFILE%\Outlook"
) do (
  if not defined DESTDIR if exist "%%~D\src\index.js" set "DESTDIR=%%~D"
)

if not defined DESTDIR (
  echo.
  echo   Outlook folder NOT found on this PC.
  echo   1. Ask the admin for the "Outlook" folder and copy it to your PC.
  echo   2. Install Node.js, then run "npm install" and "npm run login" in that folder once.
  echo   3. Run this file again.
  echo.
  pause
  exit /b 1
)

set "DEST=%DESTDIR%\OUTLOOK_bar.ahk"

echo.
echo   Installing / updating the OUTLOOK Bar...
echo   Folder: %DESTDIR%
echo.

rem  Download the latest bar (cache-busted)
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -UseBasicParsing -Uri ('%URL%?v=' + (Get-Random)) -OutFile '%DEST%'; exit 0 } catch { exit 1 }"
if errorlevel 1 (
  echo.
  echo   Download FAILED. Check your internet connection and run this again.
  echo.
  pause
  exit /b 1
)

rem  Point the desktop shortcut at the bar
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $l=$w.CreateShortcut([Environment]::GetFolderPath('Desktop')+'\OUTLOOK Bar.lnk'); $l.TargetPath='%DEST%'; $l.WorkingDirectory='%DESTDIR%'; $l.Save()"

echo   Done. Starting the bar...
start "" "%DEST%"

timeout /t 3 >nul
exit /b 0
