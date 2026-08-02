@echo off
setlocal EnableExtensions
chcp 65001 >nul
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
  "%USERPROFILE%\Documents\Outlook"
  "%USERPROFILE%\GitHub\Outlook"
  "%USERPROFILE%\Outlook"
  "%USERPROFILE%\Desktop\Outlook"
  "%USERPROFILE%\Downloads\Outlook"
) do (
  if not defined DESTDIR if exist "%%~D\src\index.js" set "DESTDIR=%%~D"
)
if not defined DESTDIR goto :nofolder

set "DEST=%DESTDIR%\OUTLOOK_bar.ahk"
echo.
echo   OUTLOOK Bar - install / update
echo   Folder: %DESTDIR%
echo.

rem  ---- [1/4] Node.js ----
where node >nul 2>nul
if not errorlevel 1 goto :havenode
if exist "%ProgramFiles%\nodejs\node.exe" (
  set "PATH=%ProgramFiles%\nodejs;%PATH%"
  goto :havenode
)
echo   [1/4] Installing Node.js... (one time - please wait / 처음 한 번 - 잠시만요)
where winget >nul 2>nul
if errorlevel 1 goto :nodefail
winget install --id OpenJS.NodeJS.LTS -e --accept-source-agreements --accept-package-agreements
if exist "%ProgramFiles%\nodejs\node.exe" (
  set "PATH=%ProgramFiles%\nodejs;%PATH%"
  goto :havenode
)
where node >nul 2>nul
if not errorlevel 1 goto :havenode
goto :nodefail
:havenode

rem  ---- [2/4] npm install (first time only) ----
if exist "%DESTDIR%\node_modules" goto :havedeps
echo   [2/4] Preparing files... (one time, 1-3 min / 처음 한 번, 1~3분)
pushd "%DESTDIR%"
call npm install
popd
:havedeps

rem  ---- [3/4] login (first time only) ----
if exist "%DESTDIR%\tokens.json" goto :havelogin
echo.
echo   [3/4] A LOGIN window will open - sign in with your WORK email.
echo         로그인 창이 뜨면 "회사 메일"로 로그인하세요.
echo.
pushd "%DESTDIR%"
call npm run login
popd
:havelogin

rem  ---- [4/4] download the latest bar (cache-busted) ----
echo   [4/4] Getting the latest bar...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -UseBasicParsing -Uri ('%URL%?v=' + (Get-Random)) -OutFile '%DEST%'; exit 0 } catch { exit 1 }"
if errorlevel 1 goto :dlfail

rem  Desktop shortcut
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $l=$w.CreateShortcut([Environment]::GetFolderPath('Desktop')+'\OUTLOOK Bar.lnk'); $l.TargetPath='%DEST%'; $l.WorkingDirectory='%DESTDIR%'; $l.Save()"

echo.
echo   Done! Starting the bar... / 완료! 바를 켭니다...
start "" "%DEST%"
timeout /t 3 >nul
exit /b 0

:nofolder
echo.
echo   "Outlook" folder NOT found on this PC.
echo   Ask the admin for the Outlook folder, copy it to your Documents folder,
echo   then run this file again.
echo.
echo   이 PC에 "Outlook" 폴더가 없어요.
echo   관리자에게 Outlook 폴더를 받아 문서(Documents) 폴더에 넣고,
echo   이 파일을 다시 실행하세요.
echo.
pause
exit /b 1

:nodefail
echo.
echo   Could not install Node.js automatically.
echo   A download page will open - install it, then run this file again.
echo   Node.js 자동 설치에 실패했어요. 열리는 페이지에서 설치 후 다시 실행하세요.
echo.
start https://nodejs.org
pause
exit /b 1

:dlfail
echo.
echo   Download FAILED. Check your internet connection and run this again.
echo   다운로드 실패 - 인터넷 확인 후 다시 실행하세요.
echo.
pause
exit /b 1
