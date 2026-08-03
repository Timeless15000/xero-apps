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
  "%USERPROFILE%\OUTLOOK Bar\Outlook"
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
rem  Managed staff copy: always refresh it from the company folder (auto-update)
if /i "%DESTDIR%"=="%USERPROFILE%\OUTLOOK Bar\Outlook" goto :findshared
rem  Also check the REAL Documents/Desktop folders (OneDrive-redirected PCs)
if defined DESTDIR goto :havefolder
for /f "usebackq delims=" %%P in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('MyDocuments'); [Environment]::GetFolderPath('Desktop'); [Environment]::GetFolderPath('UserProfile')+'\Downloads'"`) do (
  if not defined DESTDIR if exist "%%P\Outlook\src\index.js" set "DESTDIR=%%P\Outlook"
)
if defined DESTDIR goto :havefolder

rem  Locate the company shared copy (works for any OneDrive/SharePoint layout)
:findshared
set "SRCDIR="
for /f "usebackq delims=" %%S in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "$mids=@('Timeless 042026 - Documents\Admin\Automation\Outlook','Admin\Automation\Outlook','Automation\Outlook','Timeless 042026 - Documents\Admin\Automation\OUTLOOK Bar\Outlook'); $roots=@($env:OneDriveCommercial,$env:OneDrive); Get-ChildItem $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | ForEach-Object { $roots += $_.FullName }; foreach($r in $roots){ if(-not $r){continue}; foreach($m in $mids){ $p=Join-Path $r $m; if(Test-Path (Join-Path $p 'src\index.js')){ Write-Output $p; exit } } }"`) do if not defined SRCDIR set "SRCDIR=%%S"
if not defined SRCDIR goto :nofolder
echo   Getting the latest program from the company folder... / 회사 폴더에서 최신 프로그램 가져오는 중...
set "DESTDIR=%USERPROFILE%\OUTLOOK Bar\Outlook"
robocopy "%SRCDIR%" "%DESTDIR%" /E /XD node_modules reports .git .claude docs /XF tokens.json state.json log.txt "flagged-cache*.json" OUTLOOK_bar.ini >nul
if not exist "%DESTDIR%\src\index.js" goto :nofolder
:havefolder

set "DEST=%DESTDIR%\OUTLOOK_bar.ahk"
echo.
echo   OUTLOOK Bar - install / update
echo   Folder: %DESTDIR%
echo.

rem  ---- [1/3] Node.js (fully automatic - no admin rights needed) ----
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
echo   [1/3] Installing Node.js... please wait 1-3 min, do NOT close this window
echo          Node.js 설치 중... 1~3분 걸려요. 창을 닫지 마세요
set "NARCH=win-x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "NARCH=win-arm64"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; $ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $b='https://nodejs.org/dist/latest-v24.x/'; $s=(Invoke-WebRequest -UseBasicParsing ($b+'SHASUMS256.txt')).Content; $n=[regex]::Match($s,'node-v[0-9.]+-%NARCH%\.zip').Value; if(-not $n){exit 1}; $z=Join-Path $env:TEMP 'node-lts.zip'; Invoke-WebRequest -UseBasicParsing ($b+$n) -OutFile $z; $t=Join-Path $env:TEMP 'node-lts-unzip'; if(Test-Path $t){Remove-Item $t -Recurse -Force}; Expand-Archive -Path $z -DestinationPath $t -Force; $i=Get-ChildItem $t -Directory | Select-Object -First 1; $d='%NODEDIR%'; if(Test-Path $d){Remove-Item $d -Recurse -Force}; Move-Item $i.FullName $d; Remove-Item $z -Force; Remove-Item $t -Recurse -Force"
if exist "%NODEDIR%\node.exe" (
  set "PATH=%NODEDIR%;%PATH%"
  goto :havenode
)
goto :nodefail
:havenode
rem  Make portable Node available to the bar later too (user PATH, one time)
if exist "%NODEDIR%\node.exe" powershell -NoProfile -ExecutionPolicy Bypass -Command "$d='%NODEDIR%'; $p=[Environment]::GetEnvironmentVariable('Path','User'); if(-not $p){$p=''}; if($p -notlike ('*'+$d+'*')){[Environment]::SetEnvironmentVariable('Path',($p.TrimEnd(';')+';'+$d),'User')}"

rem  ---- [2/3] npm install (first time only) ----
if exist "%DESTDIR%\node_modules" goto :havedeps
echo   [2/3] Preparing files... (one time, 1-3 min / 처음 한 번, 1~3분)
pushd "%DESTDIR%"
call npm install
popd
:havedeps


rem  ---- AutoHotkey v2 (auto install if missing - no admin needed) ----
call :findahk
if defined AHKEXE goto :haveahk
echo   Installing AutoHotkey... (one time, 1 min / 자동 설치 - 처음 한 번, 1분)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -UseBasicParsing 'https://www.autohotkey.com/download/ahk-v2.exe' -OutFile ($env:TEMP+'\ahk-v2-setup.exe'); exit 0 } catch { exit 1 }"
if not exist "%TEMP%\ahk-v2-setup.exe" goto :ahkfail
start /wait "" "%TEMP%\ahk-v2-setup.exe" /silent
call :findahk
if not defined AHKEXE goto :ahkfail
:haveahk

rem  ---- [3/3] download the latest bar (cache-busted) ----
echo   [3/3] Getting the latest bar...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -UseBasicParsing -Uri ('%URL%?v=' + (Get-Random)) -OutFile '%DEST%'; exit 0 } catch { exit 1 }"
if errorlevel 1 goto :dlfail

rem  Outlook logo for the desktop shortcut (best effort)
if not exist "%DESTDIR%\outlook.ico" powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; try { [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri ('%ICOURL%?v=' + (Get-Random)) -OutFile '%DESTDIR%\outlook.ico' } catch {}"

rem  Desktop shortcut (with Outlook icon when available)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $l=$w.CreateShortcut([Environment]::GetFolderPath('Desktop')+'\OUTLOOK Bar.lnk'); $l.TargetPath='%DEST%'; $l.WorkingDirectory='%DESTDIR%'; if(Test-Path '%DESTDIR%\outlook.ico'){$l.IconLocation='%DESTDIR%\outlook.ico,0'}; $l.Save()"
ie4uinit.exe -show >nul 2>nul

echo.
echo   Done! Starting the bar... / 완료! 바를 켭니다...
start "" "%AHKEXE%" "%DEST%"
timeout /t 3 >nul
exit /b 0

:findahk
set "AHKEXE="
if exist "%LocalAppData%\Programs\AutoHotkey\v2\AutoHotkey64.exe" set "AHKEXE=%LocalAppData%\Programs\AutoHotkey\v2\AutoHotkey64.exe"
if not defined AHKEXE if exist "%ProgramFiles%\AutoHotkey\v2\AutoHotkey64.exe" set "AHKEXE=%ProgramFiles%\AutoHotkey\v2\AutoHotkey64.exe"
if not defined AHKEXE if exist "%LocalAppData%\Programs\AutoHotkey\v2\AutoHotkey32.exe" set "AHKEXE=%LocalAppData%\Programs\AutoHotkey\v2\AutoHotkey32.exe"
if not defined AHKEXE if exist "%ProgramFiles%\AutoHotkey\v2\AutoHotkey32.exe" set "AHKEXE=%ProgramFiles%\AutoHotkey\v2\AutoHotkey32.exe"
exit /b 0

:ahkfail
echo.
echo   Could not install AutoHotkey automatically.
echo   A download page will open - install v2, then run this file again.
echo   AutoHotkey 자동 설치에 실패했어요. 열리는 페이지에서 v2 설치 후 다시 실행하세요.
echo.
start https://www.autohotkey.com
pause
exit /b 1

:nofolder
echo.
echo   OUTLOOK program not found on this PC yet.
echo   Make sure the company OneDrive is signed in and synced,
echo   wait a few minutes, then run this file again.
echo   If it still fails, contact the admin.
echo.
echo   이 PC에서 OUTLOOK 프로그램을 아직 못 찾았어요.
echo   회사 OneDrive 로그인/동기화가 됐는지 확인하고,
echo   몇 분 뒤에 이 파일을 다시 실행하세요.
echo   계속 안 되면 관리자(Brian)에게 알려주세요.
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
