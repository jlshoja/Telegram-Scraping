@echo off
chcp 65001 >nul
cd /d "%~dp0"

:menu
echo.
echo  ========================================
echo   Telegram Scraping - Runner
echo  ========================================
echo   1. Fill Excel (cached posts)
echo   2. Rescan channels, then fill
echo   3. Fix categories (standardize category)
echo   4. Telegram login (first time only)
echo   5. Build WooCommerce import sheet
echo   6. Download product images
echo   0. Exit
echo  ========================================
if not "%~1"=="" (set "opt=%~1" & goto :run)
set /p opt="Select option: "
set "opt=%opt: =%"
if "%opt%"=="" (echo No input, exiting. & exit /b 1)

:run

if "%opt%"=="1" python fill.py & goto :end
if "%opt%"=="2" python scrape.py & python fill.py & goto :end
if "%opt%"=="3" python fill.py --categories-only & goto :end
if "%opt%"=="4" goto :login
if "%opt%"=="5" python build_woocommerce.py & goto :end
if "%opt%"=="6" goto :download
if "%opt%"=="0" exit /b 0
echo Invalid option, try again.
goto :menu

:download
python download_images.py
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
