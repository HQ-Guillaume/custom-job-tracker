@echo off
setlocal
cd /d "%~dp0"

title Custom Job Tracker GUI
echo Starting Custom Job Tracker GUI...
echo.
echo This fallback window stays attached to the GUI so startup errors remain visible.
echo Use Run-CustomJobTracker-GUI.vbs when you want the no-console launcher.
echo.

powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0app\cli\Launch-AnalyticsJobCrawlerGui.ps1"
set "GUI_EXIT_CODE=%ERRORLEVEL%"

if not "%GUI_EXIT_CODE%"=="0" (
    echo.
    echo GUI closed with error code %GUI_EXIT_CODE%.
    echo Check output\launcher_error.log if it exists.
    echo.
    pause
)

exit /b %GUI_EXIT_CODE%
