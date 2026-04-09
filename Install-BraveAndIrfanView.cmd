@echo off
setlocal enabledelayedexpansion

echo Checking for installed applications...

REM Detect system architecture for ARM64 installer preference
set "IsArm64=false"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "IsArm64=true"
if /i "%PROCESSOR_ARCHITEW6432%"=="ARM64" set "IsArm64=true"
if "%IsArm64%"=="true" (
    echo ARM64 system detected. ARM installers will be preferred when available.
)

REM Check if Brave is installed
set "BraveInstalled=false"
if exist "%ProgramFiles%\BraveSoftware\Brave-Browser\Application\brave.exe" (
    set "BraveInstalled=true"
    echo Brave is already installed.
)
if exist "%ProgramFiles(x86)%\BraveSoftware\Brave-Browser\Application\brave.exe" (
    set "BraveInstalled=true"
    echo Brave is already installed.
)

REM Check if IrfanView is installed
set "IrfanViewInstalled=false"
if exist "%ProgramFiles%\IrfanView\i_view*.exe" (
    set "IrfanViewInstalled=true"
    echo IrfanView is already installed.
)
if exist "%ProgramFiles(x86)%\IrfanView\i_view*.exe" (
    set "IrfanViewInstalled=true"
    echo IrfanView is already installed.
)

REM Check for ISO with WinApps directory
set "ISOFound=false"
set "WinAppsDir="

REM Create a temporary PowerShell script to find ISO drives
echo $isoDrives = Get-Volume ^| Where-Object { $_.DriveType -eq 'CD-ROM' -and $_.OperationalStatus -eq 'OK' } ^| Select-Object -ExpandProperty DriveLetter > "%TEMP%\find_iso.ps1"
echo foreach($drive in $isoDrives) { >> "%TEMP%\find_iso.ps1"
echo   if (Test-Path "$drive`:\WinApps") { >> "%TEMP%\find_iso.ps1"
echo     Write-Host "$drive" >> "%TEMP%\find_iso.ps1"
echo     exit >> "%TEMP%\find_iso.ps1"
echo   } >> "%TEMP%\find_iso.ps1"
echo } >> "%TEMP%\find_iso.ps1"

REM Run the PowerShell script and capture its output
for /f %%i in ('powershell -ExecutionPolicy Bypass -File "%TEMP%\find_iso.ps1"') do (
    set "IsoDrive=%%i"
    set "ISOFound=true"
    set "WinAppsDir=%%i:\WinApps"
    echo Found WinApps directory on drive %%i:
)

REM Clean up the temporary script
del "%TEMP%\find_iso.ps1" > nul 2>&1

REM Install Brave if not already installed
if "%BraveInstalled%"=="false" (
    echo Installing Brave Browser...
    
    if "%ISOFound%"=="true" (
        if "%IsArm64%"=="true" (
            if exist "%WinAppsDir%\brave_silent_arm.exe" (
                echo Installing Brave ARM version from ISO...
                start /wait "" "%WinAppsDir%\brave_silent_arm.exe"
                echo Brave ARM installation from ISO completed.
            ) else if exist "%WinAppsDir%\brave_silent.exe" (
                echo ARM Brave installer not found, installing Brave from ISO fallback...
                start /wait "" "%WinAppsDir%\brave_silent.exe"
                echo Brave installation from ISO fallback completed.
            ) else (
                echo Brave installer not found on ISO, installing from online source...
                winget install Brave.Brave --accept-package-agreements --accept-source-agreements
            )
        ) else (
            if exist "%WinAppsDir%\brave_silent.exe" (
                echo Installing Brave from ISO...
                start /wait "" "%WinAppsDir%\brave_silent.exe"
                echo Brave installation from ISO completed.
            ) else (
                echo Brave installer not found on ISO, installing from online source...
                winget install Brave.Brave --accept-package-agreements --accept-source-agreements
            )
        )
    ) else (
        echo Installing Brave from online source...
        winget install Brave.Brave --accept-package-agreements --accept-source-agreements
    )
) else (
    echo Skipping Brave installation as it's already installed.
)

REM Install IrfanView if not already installed
if "%IrfanViewInstalled%"=="false" (
    echo Installing IrfanView...
    
    if "%ISOFound%"=="true" (
        if "%IsArm64%"=="true" (
            if exist "%WinAppsDir%\iview_setup_arm64.exe" (
                echo Installing IrfanView ARM64 version from ISO...
                start /wait "" "%WinAppsDir%\iview_setup_arm64.exe" /silent
                echo IrfanView ARM64 installation from ISO completed.
            ) else if exist "%WinAppsDir%\iview_setup.exe" (
                echo ARM64 IrfanView installer not found, installing IrfanView from ISO fallback...
                start /wait "" "%WinAppsDir%\iview_setup.exe" /silent
                echo IrfanView installation from ISO fallback completed.
            ) else (
                echo IrfanView installer not found on ISO, installing from online source...
                winget install IrfanSkiljan.IrfanView --accept-package-agreements --accept-source-agreements
            )
        ) else (
            if exist "%WinAppsDir%\iview_setup.exe" (
                echo Installing IrfanView from ISO...
                start /wait "" "%WinAppsDir%\iview_setup.exe" /silent
                echo IrfanView installation from ISO completed.
            ) else (
                echo IrfanView installer not found on ISO, installing from online source...
                winget install IrfanSkiljan.IrfanView --accept-package-agreements --accept-source-agreements
            )
        )
    ) else (
        echo Installing IrfanView from online source...
        winget install IrfanSkiljan.IrfanView --accept-package-agreements --accept-source-agreements
    )
) else (
    echo Skipping IrfanView installation as it's already installed.
)

echo Installation process completed.
exit /b 0