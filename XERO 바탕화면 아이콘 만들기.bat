@echo off
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\XERO Bar.lnk'); $sc.TargetPath = '%~dp0XERO_bar.ahk'; $sc.WorkingDirectory = '%~dp0'; $sc.IconLocation = '%~dp0XERO_bar.ahk,0'; $sc.Save()"
echo.
echo   Done. A "XERO Bar" icon was created on your Desktop.
echo   Double-click that icon to open the bar.
echo.
pause
