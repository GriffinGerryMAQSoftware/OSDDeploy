
@echo off
set LOG=C:\Windows\Temp\SetupComplete-OSDCloud.log
set OSDROOT=C:\ProgramData\OSDCloud
set ASSETS=%OSDROOT%\Assets
set SCRIPTS=%OSDROOT%\Scripts

mkdir "%ASSETS%" 2>nul
mkdir "%SCRIPTS%" 2>nul

echo ==== SetupComplete start ==== >> "%LOG%"

:: Copy MAQAuditor MSI from USB to local disk
echo Copying MAQAuditor.msi >> "%LOG%"
copy "X:\OSDCloud\Config\Assets\MAQAuditor.msi" "%ASSETS%\MAQAuditor.msi" /Y >> "%LOG%" 2>&1

:: Copy wallpaper from USB to local disk
echo Copying MAQSoftware.jpg >> "%LOG%"
copy "X:\OSDCloud\Config\Assets\MAQSoftware.jpg" "%ASSETS%\MAQSoftware.jpg" /Y >> "%LOG%" 2>&1

:: Copy OOBE script locally
echo Copying oobe.ps1 >> "%LOG%"
copy "X:\OSDCloud\Config\Scripts\oobe.ps1" "%SCRIPTS%\oobe.ps1" /Y >> "%LOG%" 2>&1

:: Sanity checks
if not exist "%ASSETS%\MAQAuditor.msi" (
  echo ERROR: MAQAuditor.msi missing >> "%LOG%"
  exit /b 1
)

if not exist "%ASSETS%\MAQSoftware.jpg" (
  echo ERROR: MAQSoftware.jpg missing >> "%LOG%"
  exit /b 1
)

if not exist "%SCRIPTS%\oobe.ps1" (
  echo ERROR: oobe.ps1 missing >> "%LOG%"
  exit /b 1
)

:: Execute OOBE tasks from local disk
echo Executing oobe.ps1 >> "%LOG%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTS%\oobe.ps1" >> "%LOG%" 2>&1

echo ==== SetupComplete end ==== >> "%LOG%"
exit /b 0
