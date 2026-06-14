@echo off
setlocal
cd /d "%~dp0"

title Analytics Job Crawler - Starting
cls
echo Analytics Job Crawler
echo Started: %DATE% %TIME%
echo.
echo The window title and timestamped lines show the current stage.
echo The crawl can take a few minutes, especially during LinkedIn detail fetches.
echo Main file: output\jobs_tracker.xlsx
echo.
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Find-AnalyticsJobs.ps1"
set "CRAWLER_EXIT_CODE=%ERRORLEVEL%"

echo.
if "%CRAWLER_EXIT_CODE%"=="0" (
    title Analytics Job Crawler - Finished
    echo Finished successfully.
) else (
    title Analytics Job Crawler - Error
    echo Finished with error code %CRAWLER_EXIT_CODE%.
)
echo Tracker workbook and backups are in:
echo %~dp0output
echo.
pause
exit /b %CRAWLER_EXIT_CODE%
