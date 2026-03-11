# FirstLogon.ps1 (USER)
# Purpose: user-context installs (WinGet, Office ODT) after first sign-in.

$ErrorActionPreference = 'Continue'

$Root   = 'C:\ProgramData\OSDCloud'
$Assets = Join-Path $Root 'Assets'
$Logs   = Join-Path $Root 'Logs'
$Flags  = Join-Path $Root 'Flags'

New-Item -ItemType Directory -Force -Path $Logs,$Flags | Out-Null

$LogFile = Join-Path $Logs 'FirstLogon.log'
Start-Transcript -Path (Join-Path $Logs 'FirstLogon-transcript.log') -Append | Out-Null

function Write-Log([string]$m){
  Add-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $m"
}

$DoneFlag = Join-Path $Flags '.firstlogon.done'
if (Test-Path $DoneFlag) {
  Write-Log "Done flag exists; exiting."
  exit 0
}

try {
  Write-Log "=== FirstLogon start (User=$env:USERNAME) ==="

  # WinGet may not be available until first logon triggers registration
  # Microsoft guidance suggests requesting registration via DesktopAppInstaller. [3](https://akosbakos.ch/osdcloud-9-oobe-challenges/)
  try {
    Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction SilentlyContinue
    Write-Log "Requested DesktopAppInstaller registration for WinGet."
  } catch {
    Write-Log "WinGet registration request failed: $($_.Exception.Message)"
  }

  function Get-WingetPath {
    $cmd = Get-Command winget -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $c = Get-ChildItem "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue |
         Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($c) { return $c.FullName }

    return $null
  }

  $winget = Get-WingetPath
  if (-not $winget) {
    Write-Log "WinGet still not available. Leaving task to retry next logon."
    exit 0
  }

  Write-Log "WinGet resolved: $winget"

  $apps = @(
    @{Name='VS Code Insiders';     Id='Microsoft.VisualStudioCode.Insiders'},
    @{Name='Dell Command Update';  Id='dell.commandupdate'},
    @{Name='Adobe Acrobat Reader'; Id='Adobe.Acrobat.Reader.64-bit'},
    @{Name='Google Chrome';        Id='Google.Chrome'}
  )

  foreach ($app in $apps) {
    Write-Log "Installing $($app.Name) ($($app.Id))..."
    & $winget install -e --silent --disable-interactivity --accept-package-agreements --accept-source-agreements --source winget --id $app.Id 2>&1 |
      ForEach-Object { Write-Log $_ }
    Write-Log "WinGet ExitCode=$LASTEXITCODE for $($app.Name)"
    Start-Sleep -Seconds 2
  }

  # Optional Office ODT install (offline payload staged)
  $odtSetup = Join-Path $Assets 'Office\setup.exe'
  $cfg1     = Join-Path $Assets 'Office\config.xml'

  if (Test-Path $odtSetup -and Test-Path $cfg1) {
    Write-Log "Installing Office ODT using $cfg1"
    $p = Start-Process -FilePath $odtSetup -ArgumentList "/configure `"$cfg1`"" -Wait -PassThru
    Write-Log "Office ODT finished. ExitCode=$($p.ExitCode)"
  } else {
    Write-Log "Office ODT not installed (setup.exe or config.xml missing)."
  }

  # Mark done + delete task so it doesn't run every logon
  New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
  schtasks /Delete /TN "OSDCloud-FirstLogon" /F | Out-Null

  Write-Log "=== FirstLogon completed ==="
}
finally {
  Stop-Transcript | Out-Null
}