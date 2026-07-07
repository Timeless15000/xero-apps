@echo off
setlocal enabledelayedexpansion
title Fix and Restore XERO apps
cd /d "C:\Users\dudeo\OneDrive\Document\GitHub\xero-apps"

REM --- locate git (PATH first, else GitHub Desktop's bundled git) ---
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
echo   Fix and Restore XERO apps
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

echo [1/6] Clearing stuck lock and rebuilding index...
if exist ".git\index.lock" del /f /q ".git\index.lock"
if exist ".git\index" del /f /q ".git\index"
if exist ".git\gc.log" del /f /q ".git\gc.log"
"!GIT!" reset

echo [2/6] Restoring master Xero_applications.html (v37)...
"!GIT!" checkout 27c628c -- Xero_applications.html

echo [3/6] Removing the stray duplicate...
if exist ".github\workflows\Xero_applications.html" del /f /q ".github\workflows\Xero_applications.html"

echo [4/6] Saving all changes...
"!GIT!" add -A
"!GIT!" commit -m "restore master Xero_applications.html v37 + index/404 redirects"

echo [5/6] Syncing with GitHub...
"!GIT!" pull --no-edit

echo [6/6] Uploading...
"!GIT!" push

echo.
echo ============================================
echo   DONE. Wait about 1-2 minutes, then open
echo   the XERO apps page and press Ctrl+Shift+R.
echo ============================================
echo.
pause
