@echo off
setlocal EnableExtensions
chcp 65001 >nul
title XERO Bar - install / update

set "DESTDIR=%USERPROFILE%\XERO Bar"
set "DEST=%DESTDIR%\XERO_bar.ahk"
set "URL=https://raw.githubusercontent.com/Timeless15000/xero-apps/main/XERO_bar.ahk"
set "ICOURL=https://raw.githubusercontent.com/Timeless15000/xero-apps/main/xero.ico"

echo.
echo   Installing / updating the XERO Bar...
echo.

if not exist "%DESTDIR%" mkdir "%DESTDIR%"

rem  ---- AutoHotkey v2 (auto install if missing - no admin needed) ----
call :findahk
if defined AHKEXE goto :haveahk
echo   Installing AutoHotkey... (one time, 1 min / 자동 설치 - 처음 한 번, 1분)
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -UseBasicParsing 'https://www.autohotkey.com/download/ahk-v2.exe' -OutFile ($env:TEMP+'\ahk-v2-setup.exe'); exit 0 } catch { exit 1 }"
if not exist "%TEMP%\ahk-v2-setup.exe" goto :ahkfail
start /wait "" "%TEMP%\ahk-v2-setup.exe" /silent
call :findahk
if not defined AHKEXE goto :ahkfail
:haveahk

rem  Download the latest bar (cache-busted)
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -UseBasicParsing -Uri ('%URL%?v=' + (Get-Random)) -OutFile '%DEST%'; exit 0 } catch { exit 1 }"
if errorlevel 1 goto :dlfail

rem  Xero logo for the desktop shortcut (best effort)
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri ('%ICOURL%?v=' + (Get-Random)) -OutFile '%DESTDIR%\xero.ico' } catch {}"

rem  Point the desktop shortcut at the new self-updating bar (with Xero icon when available)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $l=$w.CreateShortcut([Environment]::GetFolderPath('Desktop')+'\XERO Bar.lnk'); $l.TargetPath='%DEST%'; $l.WorkingDirectory='%DESTDIR%'; if(Test-Path '%DESTDIR%\xero.ico'){$l.IconLocation='%DESTDIR%\xero.ico,0'}; $l.Save()"
ie4uinit.exe -show >nul 2>nul

echo   Done. Starting the bar... / 완료! 바를 켭니다...
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

:dlfail
echo.
echo   Download FAILED. Check your internet connection and run this again.
echo   다운로드 실패 - 인터넷 확인 후 다시 실행하세요.
echo.
pause
exit /b 1
