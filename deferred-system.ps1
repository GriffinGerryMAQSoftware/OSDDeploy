# Deferred-System.ps1 (SYSTEM)
# Purpose: silent machine installs/config after setup without blocking SetupComplete UI.

$ErrorActionPreference = 'Continue'

$Root   = 'C:\ProgramData\OSDCloud'
$Assets = Join-Path $Root 'Assets'
$Logs   = Join-Path $Root 'Logs'
$Flags  = Join-Path $Root 'Flags'

New-Item -ItemType Directory -Force -Path $Logs,$Flags | Out-Null

$LogFile = Join-Path $Logs 'Deferred-System.log'
Start-Transcript -Path (Join-Path $Logs 'Deferred-System-transcript.log') -Append | Out-Null

function Write-Log([string]$m){
  Add-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $m"
}

$DoneFlag = Join-Path $Flags '.deferredsystem.done'
if (Test-Path $DoneFlag) {
  Write-Log "Done flag exists; exiting."
  exit 0
}

try {
  Write-Log "=== Deferred-System start (IsSystem=$([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem)) ==="

  # Give services/network a moment after boot
  Start-Sleep -Seconds 60

  # Install MAQAuditor MSI silently
  $MsiPath = Join-Path $Assets 'MAQAuditor.msi'
  $MsiLog  = Join-Path $Assets 'MAQAuditor_install.log'

  if (Test-Path $MsiPath) {
    Write-Log "Installing MAQAuditor MSI: $MsiPath"
    $args = "/i `"$MsiPath`" /qn /norestart /L*v `"$MsiLog`""

    $p = Start-Process msiexec.exe -ArgumentList $args -PassThru
    $timeoutSec = 1200

    if (Wait-Process -Id $p.Id -Timeout $timeoutSec -ErrorAction SilentlyContinue) {
      Write-Log "MAQAuditor install finished. ExitCode=$($p.ExitCode)"
    } else {
      Write-Log "WARNING: MAQAuditor install exceeded ${timeoutSec}s; continuing."
      try { Stop-Process -Id $p.Id -Force } catch {}
    }
  } else {
    Write-Log "WARNING: MAQAuditor MSI not found at $MsiPath"
  }

  # Mark done + delete task
  New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
  schtasks /Delete /TN "OSDCloud-DeferredSystem" /F | Out-Null

  Write-Log "=== Deferred-System completed ==="
}
finally {
  Stop-Transcript | Out-Null
}