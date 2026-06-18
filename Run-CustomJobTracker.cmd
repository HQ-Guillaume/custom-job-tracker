@echo off
setlocal
cd /d "%~dp0"

title Custom Job Tracker - Starting
cls
echo Custom Job Tracker
echo Started: %DATE% %TIME%
echo.
echo The window title and timestamped lines show the current stage.
echo The crawl can take a few minutes, especially during LinkedIn detail fetches.
echo Main file: output\jobs_tracker.xlsx
echo.

set "CRAWL_MODE=%~1"
if not "%CRAWL_MODE%"=="" goto mode_selected

echo Choose crawl mode:
echo   [F] Fast    - quicker, lighter coverage
echo   [D] Default - balanced daily crawl
echo   [P] Deep    - slower, maximum coverage
echo.
choice /C FDP /N /M "Mode [F/D/P]: "
if errorlevel 3 set "CRAWL_MODE=Deep"
if errorlevel 2 set "CRAWL_MODE=Default"
if errorlevel 1 set "CRAWL_MODE=Fast"

:mode_selected
echo.
echo Selected mode: %CRAWL_MODE%
echo.
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\cli\Find-AnalyticsJobs.ps1" -CrawlMode "%CRAWL_MODE%"
set "CRAWLER_EXIT_CODE=%ERRORLEVEL%"

echo.
if "%CRAWLER_EXIT_CODE%"=="0" (
    title Custom Job Tracker - Finished
    echo Finished successfully.
) else (
    title Custom Job Tracker - Error
    echo Finished with error code %CRAWLER_EXIT_CODE%.
)
echo Tracker workbook and backups are in:
echo %~dp0output
echo.
pause
exit /b %CRAWLER_EXIT_CODE%
