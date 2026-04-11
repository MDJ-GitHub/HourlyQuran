@echo off
setlocal enabledelayedexpansion

:: --- Request Administrator Rights ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Administrator privileges required...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

title Hourly Quran Task Manager
color 0B

:menu
cls
echo.
echo  +-+ +-+ +-+ +-+ +-+ +-+   +-+ +-+ +-+ +-+ +-+
echo  ^|H^| ^|O^| ^|U^| ^|R^| ^|L^| ^|Y^|   ^|Q^| ^|U^| ^|R^| ^|A^| ^|N^|
echo  +-+ +-+ +-+ +-+ +-+ +-+   +-+ +-+ +-+ +-+ +-+
echo.
echo  ##  ##   ####   ##  ##  ######  ##      ##  ##
echo  ##  ##  ##  ##  ##  ##  ##  ##  ##      ## ##
echo  ######  ##  ##  ##  ##  ######  ##       ###
echo  ##  ##  ##  ##  ##  ##  ####    ##       ## 
echo  ##  ##   ####    ####   ##  ##  ######  ##
echo.
echo     ####   ##  ##  ######   ####   ##   ##
echo    ##  ##  ##  ##  ##  ##  ##  ##  #### ##
echo    ##  ##  ##  ##  ######  ######  ## ####
echo    ## ###  ##  ##  ####    ##  ##  ##  ###
echo     ## ##   ####   ##  ##  ##  ##  ##   ##
echo.
echo ============================================
echo        Hourly Quran Task Manager
echo ============================================
echo.
echo   [1] Install Hourly Quran Scheduled Task
echo   [2] Uninstall Hourly Quran Scheduled Task
echo   [3] Help
echo   [0] Exit
echo.
echo ============================================
set /p choice=Choose an option (input the number then press Enter): 

if "!choice!"=="1" goto install
if "!choice!"=="2" goto uninstall
if "!choice!"=="3" goto help
if "!choice!"=="0" exit /b
goto menu

:install
cls
echo ============================================
echo Installing HourlyQuran...
echo ============================================
echo.

:: Build paths using delayed expansion to handle special chars
set "SCRIPT_PATH=%~dp0launch.vbs"

:: Check if script exists
if not exist "!SCRIPT_PATH!" (
    echo ERROR: launch.vbs not found in this folder!
    pause
    goto menu
)

:: ============================
:: CHECK AND REINSTALL TASK
:: ============================
schtasks /query /tn "HourlyQuran" >nul 2>&1
if !errorlevel!==0 (
    echo Task already exists. Reinstalling...
    schtasks /delete /tn "HourlyQuran" /f >nul 2>&1
)

echo Creating scheduled task...
schtasks /create /tn "HourlyQuran" /tr "wscript.exe \"!SCRIPT_PATH!\"" /sc hourly /mo 1 /st 00:00 /rl highest /f
if !errorlevel!==0 (
    echo SUCCESS: Task created successfully!
) else (
    echo ERROR: Failed to create task.
    pause
    goto menu
)

echo.
echo Applying battery and power settings...

:: Write PowerShell to a temp file to avoid inline quoting issues
set "PS_TEMP=%TEMP%\hq_battery.ps1"
echo $task = Get-ScheduledTask -TaskName 'HourlyQuran' > "!PS_TEMP!"
echo $s = $task.Settings >> "!PS_TEMP!"
echo $s.DisallowStartIfOnBatteries = $false >> "!PS_TEMP!"
echo $s.StopIfGoingOnBatteries = $false >> "!PS_TEMP!"
echo Set-ScheduledTask -TaskName 'HourlyQuran' -Settings $s >> "!PS_TEMP!"

powershell -ExecutionPolicy Bypass -NonInteractive -File "!PS_TEMP!" >nul 2>&1
if !errorlevel!==0 (
    echo SUCCESS: Battery mode enabled!
) else (
    echo WARNING: Battery settings could not be applied.
)
del "!PS_TEMP!" >nul 2>&1

echo.
echo ============================================
echo  You will now receive Quran verses every hour.
echo ============================================
echo.
echo IMPORTANT: Do NOT move this folder after installation.
echo If you move it, run this installer again from the new location.
echo.
pause
goto menu

:uninstall
cls
echo ============================================
echo Removing HourlyQuran scheduled task...
echo ============================================
echo.
schtasks /delete /tn "HourlyQuran" /f >nul 2>&1
if !errorlevel!==0 (
    echo SUCCESS: Task removed successfully!
) else (
    echo ERROR: Task not found or removal failed.
)
echo.
pause
goto menu

:help
cls
echo ============================================
echo                   HELP
echo ============================================
echo.
echo This installs a scheduled task that launches Hourly Quran every hour.
echo This is not just a tool - it is a companion. A quiet reminder. A steady rhythm of guidance.
echo.
echo [Features]
echo - Receive a selection of verses every hour.
echo - Customize how many verses you get each time or show entire surah.
echo - Adjust the duration of the popup to match your reading speed.
echo - Let your favorite reciter read the verses for you.
echo - Runs on both AC power and battery - no interruptions.
echo - Multiple intelligent modes:
echo     Continuous Khatma (start to end)
echo     Random verses
echo     Alternating between both
echo     Smart priority (Khatma with occasional randomness)
echo.
echo Whether you seek consistency or variety, this program gently keeps you connected
echo to the Quran throughout your day.
echo Install it... and let every hour carry meaning.
echo.
echo - Input 1 to install, 2 to uninstall.
echo - Do not install in system folders like Program Files. and reinstall after changing program location.
echo.
echo For more details and guide on how to use hte program use visit the github repository page.
echo GitHub: https://github.com/MDJ-GitHub/HourlyQuran
echo.
pause 
goto menu