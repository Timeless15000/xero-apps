@echo off
setlocal enabledelayedexpansion
title Publish XERO apps
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
echo   Publishing XERO apps to GitHub
echo ============================================
echo.
echo [1/3] Saving your changes...
"!GIT!" add -A
"!GIT!" commit -m "update via Publish %DATE% %TIME%"
echo.
echo [2/3] Syncing with GitHub...
"!GIT!" pull --no-edit
echo.
echo [3/3] Uploading...
"!GIT!" push
echo.
echo ============================================
echo   DONE.  Wait about 2 minutes, then the
echo   XERO bar auto-updates (Tampermonkey).
echo ============================================
echo.
pause
