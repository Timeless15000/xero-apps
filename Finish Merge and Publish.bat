@echo off
setlocal enabledelayedexpansion
title Finish Merge and Publish XERO apps
cd /d "C:\Users\dudeo\OneDrive\Document\GitHub\xero-apps"

REM --- locate git (PATH first, else GitHub Desktop bundled git) ---
set "GIT="
where git >nul 2>nul && set "GIT=git"
if not defined GIT (
  for /d %%D in ("%LOCALAPPDATA%\GitHubDesktop\app-*") do set "GIT=%%D\resources\app\git\cmd\git.exe"
)
if not defined GIT (
  echo.
  echo  Could not find git. Is GitHub Desktop installed?
  echo.
  pause
  exit /b 1
)

echo ============================================
echo   Finish merge ^& publish (keeps Price Check v4)
echo ============================================
echo.
echo  Please CLOSE GitHub Desktop first, then press a key.
echo.
pause

REM OneDrive locks .git files, so disable Git auto housekeeping
"!GIT!" config gc.auto 0 >nul 2>nul
"!GIT!" config gc.autoDetach false >nul 2>nul
"!GIT!" config maintenance.auto false >nul 2>nul
"!GIT!" config fetch.writeCommitGraph false >nul 2>nul
"!GIT!" config core.fscache true >nul 2>nul

echo [1/4] Clearing stuck lock...
if exist ".git\index.lock" del /f /q ".git\index.lock"
if exist ".git\gc.log" del /f /q ".git\gc.log"

echo [2/4] Committing the resolved merge...
"!GIT!" add -A
"!GIT!" commit -m "resolve merge: Price Check v4 + parseMulti auto-build"

echo [3/4] Syncing with GitHub...
"!GIT!" pull --no-edit

echo [4/4] Uploading...
"!GIT!" push

echo.
echo ============================================
echo   DONE.  Wait about 2 minutes, then the
echo   XERO bar auto-updates (Tampermonkey).
echo ============================================
echo.
pause
