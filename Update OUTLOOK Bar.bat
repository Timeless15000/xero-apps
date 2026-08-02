@echo off
setlocal EnableExtensions
chcp 65001 >nul
title OUTLOOK Bar - install / update

set "URL=https://raw.githubusercontent.com/Timeless15000/xero-apps/main/OUTLOOK_bar.ahk"
set "ICOURL=https://raw.githubusercontent.com/Timeless15000/xero-apps/main/outlook.ico"
set "NODEDIR=%LOCALAPPDATA%\node-lts"

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
  "%OneDrive%\Documents\Outlook"
  "%OneDrive%\Document\Outlook"
  "%OneDrive%\Desktop\Outlook"
  "%OneDrive%\Outlook"
) do (
  if not defined DESTDIR if exist "%%~D\src\index.js" set "DESTDIR=%%~D"
)
rem  Also check the REAL Documents/Desktop folders (OneDrive-redirected PCs)
if defined DESTDIR goto :havefolder
for /f "usebackq delims=" %%P in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('MyDocuments'); [Environment]::GetFolderPath('Desktop'); [Environment]::GetFolderPath('UserProfile')+'\Downloads'"`) do (
  if not defined DESTDIR if exist "%%P\Outlook\src\index.js" set "DESTDIR=%%P\Outlook"
)
if not defined DESTDIR goto :nofolder
:havefolder

set "DEST=%DESTDIR%\OUTLOOK_bar.ahk"
echo.
echo   OUTLOOK Bar - install / update
echo   Folder: %DESTDIR%
echo.

rem  ---- [1/4] Node.js (fully automatic - no admin rights needed) ----
where node >nul 2>nul
if not errorlevel 1 goto :havenode
if exist "%ProgramFiles%\nodejs\node.exe" (
  set "PATH=%ProgramFiles%\nodejs;%PATH%"
  goto :havenode
)
if exist "%NODEDIR%\node.exe" (
  set "PATH=%NODEDIR%;%PATH%"
  goto :havenode
)
echo   [1/4] Installing Node.js automatically... (one time, 1-2 min / 자동 설치 - 처음 한 번, 1~2분)
set "NARCH=win-x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "NARCH=win-arm64"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $b='https://nodejs.org/dist/latest-v24.x/'; $s=(Invoke-WebRequest -UseBasicParsing ($b+'SHASUMS256.txt')).Content; $n=[regex]::Match($s,'node-v[0-9.]+-%NARCH%\.zip').Value; if(-not $n){exit 1}; $z=Join-Path $env:TEMP 'node-lts.zip'; Invoke-WebRequest -UseBasicParsing ($b+$n) -OutFile $z; $t=Join-Path $env:TEMP 'node-lts-unzip'; if(Test-Path $t){Remove-Item $t -Recurse -Force}; Expand-Archive -Path $z -DestinationPath $t -Force; $i=Get-ChildItem $t -Directory | Select-Object -First 1; $d='%NODEDIR%'; if(Test-Path $d){Remove-Item $d -Recurse -Force}; Move-Item $i.FullName $d; Remove-Item $z -Force; Remove-Item $t -Recurse -Force"
if exist "%NODEDIR%\node.exe" (
  set "PATH=%NODEDIR%;%PATH%"
  goto :havenode
)
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

rem  Outlook logo for the desktop shortcut (best effort)
if not exist "%DESTDIR%\outlook.ico" powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri ('%ICOURL%?v=' + (Get-Random)) -OutFile '%DESTDIR%\outlook.ico' } catch {}"

rem  Desktop shortcut (with Outlook icon when available)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $l=$w.CreateShortcut([Environment]::GetFolderPath('Desktop')+'\OUTLOOK Bar.lnk'); $l.TargetPath='%DEST%'; $l.WorkingDirectory='%DESTDIR%'; if(Test-Path '%DESTDIR%\outlook.ico'){$l.IconLocation='%DESTDIR%\outlook.ico,0'}; $l.Save()"
ie4uinit.exe -show >nul 2>nul

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
