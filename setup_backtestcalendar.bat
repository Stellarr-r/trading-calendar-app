@echo off
setlocal enabledelayedexpansion
title Strategy Analyzer Setup

:: Configuration - Set DEV_MODE before self-update check
set "DEV_MODE=false"

:: Self-update mechanism - Skip in development mode
set "LAUNCHER_URL=https://raw.githubusercontent.com/Stellarr-r/trading-calendar-app/main/setup_backtestcalendar.bat"
set "CURRENT_LAUNCHER=%~f0"
set "TEMP_LAUNCHER=%TEMP%\setup_backtestcalendar_new.bat"

if /i "%DEV_MODE%"=="true" (
    echo Development mode enabled - skipping launcher updates
    goto skip_launcher_update
)

echo Checking for launcher updates...
powershell -Command "try { $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri '%LAUNCHER_URL%' -OutFile '%TEMP_LAUNCHER%' -UseBasicParsing; exit 0 } catch { exit 1 }" >nul 2>&1

if exist "%TEMP_LAUNCHER%" (
    echo Comparing launcher versions...
    fc /B "%CURRENT_LAUNCHER%" "%TEMP_LAUNCHER%" >nul 2>&1
    if errorlevel 1 (
        echo Launcher update available - installing new version...
        echo.
        echo ================================================================================
        echo  LAUNCHER UPDATE FOUND
        echo.
        echo  A newer version of the Strategy Analyzer launcher is available.
        echo  The launcher will now update itself and restart automatically.
        echo.
        echo  Current location: %CURRENT_LAUNCHER%
        echo  This window will close and reopen with the updated launcher.
        echo ================================================================================
        echo.
        timeout /t 3 /nobreak >nul
        
        :: Create update script that will replace the current launcher
        echo @echo off > "%TEMP%\update_launcher.bat"
        echo timeout /t 2 /nobreak ^>nul >> "%TEMP%\update_launcher.bat"
        echo copy /Y "%TEMP_LAUNCHER%" "%CURRENT_LAUNCHER%" ^>nul 2^>^&1 >> "%TEMP%\update_launcher.bat"
        echo del "%TEMP_LAUNCHER%" ^>nul 2^>^&1 >> "%TEMP%\update_launcher.bat"
        echo start "" "%CURRENT_LAUNCHER%" >> "%TEMP%\update_launcher.bat"
        echo del "%%~f0" ^>nul 2^>^&1 >> "%TEMP%\update_launcher.bat"
        
        :: Launch update script and exit
        start "" "%TEMP%\update_launcher.bat"
        exit /b 0
    ) else (
        echo Launcher is up to date
        del "%TEMP_LAUNCHER%" >nul 2>&1
    )
) else (
    echo Could not check for launcher updates - continuing with current version
)

:skip_launcher_update
cls
echo.
echo ================================================================================
echo                            STRATEGY ANALYZER SETUP
echo                     Advanced Trading Analytics Platform
echo ================================================================================
echo.

:: Additional Configuration
set "DATA_DIR=%APPDATA%\StrategyAnalyzer"
set "PYTHON_FILE=%DATA_DIR%\trading_calendar.py"
set "GITHUB_URL=https://raw.githubusercontent.com/Stellarr-r/trading-calendar-app/main/trading_calendar.py"

echo [1/5] Environment Setup
echo       Application Directory: %DATA_DIR%

if not exist "%DATA_DIR%" (
    echo       Creating application directory...
    mkdir "%DATA_DIR%" 2>nul
    if exist "%DATA_DIR%" (
        echo       Directory created successfully
    ) else (
        echo       ERROR: Failed to create application directory
        echo       This may be due to insufficient permissions
        goto error_exit
    )
) else (
    echo       Application directory already exists
)

echo.
echo [2/5] Python Environment Check
python --version >nul 2>&1
if errorlevel 1 (
    echo       ERROR: Python not found in system PATH
    echo.
    echo ================================================================================
    echo  PYTHON INSTALLATION REQUIRED
    echo.
    echo  Strategy Analyzer requires Python 3.7 or newer to run.
    echo.
    echo  Download Python from: https://www.python.org/downloads/
    echo  
    echo  IMPORTANT: During installation, make sure to check the option:
    echo  "Add Python to PATH" or "Add Python to environment variables"
    echo.
    echo  After installing Python, restart this setup and try again.
    echo ================================================================================
    goto error_exit
) else (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
    echo       Found: !PYTHON_VERSION!
    echo       Python environment is ready
)

echo.
if /i "%DEV_MODE%"=="true" (
    echo [3/5] Development Mode - Local File Copy
    if exist "trading_calendar.py" (
        echo       Copying local development version...
        copy "trading_calendar.py" "%PYTHON_FILE%" >nul 2>&1
        if exist "%PYTHON_FILE%" (
            echo       Local version deployed successfully
        ) else (
            echo       ERROR: Failed to copy local version
            echo       Check file permissions in: %DATA_DIR%
            goto error_exit
        )
    ) else (
        echo       ERROR: trading_calendar.py not found in current directory
        echo       Expected location: %CD%\trading_calendar.py
        echo       Make sure you're running setup from the project directory
        goto error_exit
    )
) else (
    echo [3/5] Application Download
    echo       Repository: github.com/Stellarr-r/trading-calendar-app
    echo       Downloading latest version...
    
    powershell -Command "try { $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri '%GITHUB_URL%' -OutFile '%PYTHON_FILE%' -UseBasicParsing; exit 0 } catch { exit 1 }" >nul 2>&1
    
    if exist "%PYTHON_FILE%" (
        echo       Download completed successfully
        for %%A in ("%PYTHON_FILE%") do echo       File size: %%~zA bytes
    ) else (
        echo       WARNING: Download failed, checking for cached version...
        if exist "%PYTHON_FILE%" (
            echo       Using previously cached version
            for %%A in ("%PYTHON_FILE%") do echo       File size: %%~zA bytes
        ) else (
            echo       ERROR: No application file available
            echo.
            echo ================================================================================
            echo  DOWNLOAD FAILED
            echo.
            echo  Could not download Strategy Analyzer from GitHub repository.
            echo  
            echo  Possible causes:
            echo    * No internet connection
            echo    * Firewall blocking the download
            echo    * Antivirus software interference
            echo    * Corporate network restrictions
            echo.
            echo  Solutions:
            echo    * Check your internet connection
            echo    * Temporarily disable firewall/antivirus
            echo    * Try running as administrator
            echo ================================================================================
            goto error_exit
        )
    )
)

echo.
echo [4/5] Environment Configuration
if /i "%DEV_MODE%"=="true" (
    set "STRATEGY_ANALYZER_VERSION=DEV"
    echo       Version: DEVELOPMENT BUILD
    echo       Mode: Local development with debug features
) else (
    set "STRATEGY_ANALYZER_VERSION=1.0.3"
    echo       Version: 1.0.3 (Production Release)
    echo       Mode: Standard production build
)
set "STRATEGY_ANALYZER_DATA_DIR=%DATA_DIR%\data"
echo       Data storage: %STRATEGY_ANALYZER_DATA_DIR%
echo       Environment variables configured

echo.
echo [5/5] Application Launch
echo       Starting Strategy Analyzer...
echo       Application will open in a new window

cd /d "%DATA_DIR%"
python "%PYTHON_FILE%"

echo.
echo ================================================================================
echo  SETUP COMPLETE
echo.
echo  Strategy Analyzer has finished running.
echo  
echo  Your trading data and analysis are automatically saved to:
echo  %DATA_DIR%\data
echo.
echo  To run StellarInsight again, simply double-click this setup file.
echo ================================================================================
echo.
pause
exit /b 0

:error_exit
echo.
echo ================================================================================
echo  SETUP FAILED
echo.
echo  Strategy Analyzer setup could not complete due to the errors shown above.
echo  
echo  For technical support and troubleshooting:
echo  * Visit: github.com/Stellarr-r/trading-calendar-app/issues
echo  * Review the error messages above for specific solutions
echo  * Ensure you have administrative privileges if needed
echo.
echo  Common solutions:
echo  * Install Python with PATH option enabled
echo  * Run setup as administrator
echo  * Check internet connection for downloads
echo  * Verify antivirus/firewall settings
echo ================================================================================
echo.
pause
exit /b 1
