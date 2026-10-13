@echo off
chcp 65001 >nul
cd /d "%~dp0"

:menu
echo.
echo  ========================================
echo   Telegram Scraping - Runner
echo  ========================================
echo   1. Fill Excel (cached posts)
echo   2. Fill Excel + download media
echo   3. Rescan channels, then fill + media
echo   4. Fill specific codes (e.g. FA-1959,TR-402)
echo   5. Refresh old values with fresh data
echo   6. Telegram login (first time only)
echo   0. Exit
echo  ========================================
if not "%~1"=="" (set "opt=%~1" & set "arg1=%~2" & goto :run)
set /p opt="Select option: "
set "opt=%opt: =%"
if "%opt%"=="" (echo No input, exiting. & exit /b 1)

:run

if "%opt%"=="1" python fill.py & goto :end
if "%opt%"=="2" python fill.py --download-media & goto :end
if "%opt%"=="3" python scrape.py & python fill.py --download-media & goto :end
if "%opt%"=="4" goto :codes
if "%opt%"=="5" python fill.py --refresh & goto :end
if "%opt%"=="6" goto :login
if "%opt%"=="0" exit /b 0
echo Invalid option, try again.
goto :menu

:codes
if not "%arg1%"=="" (set codes=%arg1% & goto :runcodes)
set /p codes="Enter codes (comma-separated): "
:runcodes
python fill.py --download-media --codes %codes%
goto :end

:login
set /p phone="Phone number (e.g. +989143463191): "
python scrape.py --login --phone "%phone%" --wait-file
echo.
echo When the code arrives, paste it into cache\pending_code.txt
goto :end

:end
echo.
if not "%~1"=="" exit /b 0
pause
