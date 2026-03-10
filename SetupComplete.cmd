
@echo off
setlocal

set SRC=%~dp0
set STAGE=C:\ProgramData\OSDCloud\Assets

if not exist "%STAGE%" mkdir "%STAGE%"

REM Copy payload assets locally before any OOBE work
if exist "%SRC%Assets" (
  robocopy "%SRC%Assets" "%STAGE%" /E /R:2 /W:2 /NFL /NDL /NP
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SRC%oobe_new.ps1" -Phase SetupComplete

exit /b 0
