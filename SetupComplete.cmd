@echo off
setlocal

set SRC=%~dp0
set STAGE=C:\ProgramData\OSDCloud\Assets

if not exist "%STAGE%" mkdir "%STAGE%"

REM Stage payload assets locally (fast + safe)
if exist "%SRC%Assets" (
  robocopy "%SRC%Assets" "%STAGE%" /E /R:2 /W:2 /NFL /NDL /NP
)

REM Run SYSTEM SetupComplete staging + task scheduling
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SRC%setup.ps1"

exit /b 0