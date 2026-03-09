
@echo off
setlocal enabledelayedexpansion

set LOG=C:\Windows\Temp\SetupComplete-OSDCloud.log
set OSDROOT=C:\ProgramData\OSDCloud
set ASSETS=%OSDROOT%\Assets
set SCRIPTS=%OSDROOT%\Scripts

mkdir "%ASSETS%" 2>nul
mkdir "%SCRIPTS%" 2>nul

echo ==== SetupComplete start ==== >> "%LOG%"
echo Running as: %USERNAME% >> "%LOG%"
echo Date/Time: %DATE% %TIME% >> "%LOG%"

REM ----------------------------
REM Prefer local staged config first
REM ----------------------------
set SRC_ASSETS=C:\OSDCloud\Config\Assets
set SRC_SCRIPTS=C:\OSDCloud\Config\Scripts

if exist "%SRC_ASSETS%\MAQAuditor.msi" (
  echo Found local staged assets at %SRC_ASSETS% >> "%LOG%"
) else (
  echo Local staged assets not found. Scanning drives for OSDCloud media... >> "%LOG%"
  set SRC_ASSETS=
  set SRC_SCRIPTS=

  for %%D in (C D E F G H I J K L M N O P Q R S T U V W Y Z) do (
    if exist "%%D:\OSDCloud\Config\Assets\MAQAuditor.msi" (
      set SRC_ASSETS=%%D:\OSDCloud\Config\Assets
      set SRC_SCRIPTS=%%D:\OSDCloud\Config\Scripts
      echo Found OSDCloud media at %%D:\OSDCloud\Config >> "%LOG%"
      goto :FOUND_MEDIA
    )
  )
)

:FOUND_MEDIA
if not defined SRC_ASSETS (
  echo ERROR: Could not locate MAQAuditor.msi on local disk or any drive. >> "%LOG%"
  echo ==== SetupComplete end (FAIL) ==== >> "%LOG%"
  exit /b 1
)

REM ----------------------------
REM Copy MAQAuditor MSI
REM ----------------------------
echo Copying MAQAuditor.msi from %SRC_ASSETS% >> "%LOG%"
copy "%SRC_ASSETS%\MAQAuditor.msi" "%ASSETS%\MAQAuditor.msi" /Y >> "%LOG%" 2>&1

REM Copy wallpaper
echo Copying MAQSoftware.jpg from %SRC_ASSETS% >> "%LOG%"
copy "%SRC_ASSETS%\MAQSoftware.jpg" "%ASSETS%\MAQSoftware.jpg" /Y >> "%LOG%" 2>&1

REM Copy OOBE script
echo Copying oobe.ps1 from %SRC_SCRIPTS% >> "%LOG%"
copy "%SRC_SCRIPTS%\oobe.ps1" "%SCRIPTS%\oobe.ps1" /Y >> "%LOG%" 2>&1

REM ----------------------------
REM Sanity checks
REM ----------------------------
if not exist "%ASSETS%\MAQAuditor.msi" (
  echo ERROR: MAQAuditor.msi missing after copy >> "%LOG%"
  echo ==== SetupComplete end (FAIL) ==== >> "%LOG%"
  exit /b 1
)

if not exist "%ASSETS%\MAQSoftware.jpg" (
  echo ERROR: MAQSoftware.jpg missing after copy >> "%LOG%"
  echo ==== SetupComplete end (FAIL) ==== >> "%LOG%"
  exit /b 1
)

if not exist "%SCRIPTS%\oobe.ps1" (
  echo ERROR: oobe.ps1 missing after copy >> "%LOG%"
  echo ==== SetupComplete end (FAIL) ==== >> "%LOG%"
  exit /b 1
)

REM ----------------------------
REM Execute OOBE tasks
REM ----------------------------
echo Executing oobe.ps1 >> "%LOG%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTS%\oobe.ps1" >> "%LOG%" 2>&1

echo ==== SetupComplete end ==== >> "%LOG%"
exit /b 0
