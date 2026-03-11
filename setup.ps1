# Setup.ps1 (SYSTEM)
# Purpose: Stage payload + apply HKLM policies + schedule next-phase scripts.
# Avoid long installs here.

$ErrorActionPreference = 'Stop'

$Root   = 'C:\ProgramData\OSDCloud'
$Assets = Join-Path $Root 'Assets'
$Scripts= Join-Path $Root 'Scripts'
$Logs   = Join-Path $Root 'Logs'
$Flags  = Join-Path $Root 'Flags'

New-Item -ItemType Directory -Force -Path $Assets,$Scripts,$Logs,$Flags | Out-Null

$LogFile = Join-Path $Logs 'Setup.log'
Start-Transcript -Path (Join-Path $Logs 'Setup-transcript.log') -Append | Out-Null

function Write-Log([string]$m){
  Add-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $m"
}

try {
  Write-Log "=== Setup.ps1 start (User=$env:USERNAME, IsSystem=$([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem)) ==="
  Write-Log "PSScriptRoot=$PSScriptRoot"

  # --- Stage scripts locally ---
  foreach ($name in @('FirstLogon.ps1','Deferred-System.ps1')) {
    $src = Join-Path $PSScriptRoot $name
    if (Test-Path $src) {
      Copy-Item $src -Destination (Join-Path $Scripts $name) -Force
      Write-Log "Staged script: $name"
    } else {
      throw "Missing required script next to Setup: $src"
    }
  }

  # --- Validate required assets ---
  $msi = Join-Path $Assets 'MAQAuditor.msi'
  $jpg = Join-Path $Assets 'MAQSoftware.jpg'

  if (-not (Test-Path $msi)) { throw "Missing required asset: $msi" }
  if (-not (Test-Path $jpg)) { throw "Missing required asset: $jpg" }
  Write-Log "Assets present: MAQAuditor.msi + MAQSoftware.jpg"

  # --- Optional: Office ODT payload staging check ---
  $odtSetup = Join-Path $Assets 'Office\setup.exe'
  if (Test-Path $odtSetup) {
    Write-Log "Office ODT payload detected at: $odtSetup (will install at FirstLogon)"
  } else {
    Write-Log "Office ODT payload not found; skipping Office install scheduling."
  }

  # --- Apply HKLM baseline policies (safe in Setup) ---
  Write-Log "Applying HKLM policy baselines..."

  # Disable Consumer Apps
  New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Force | Out-Null
  New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name DisableWindowsConsumerFeatures -PropertyType DWORD -Value 1 -Force | Out-Null

  # Disable Mobile Hotspot UI
  New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Network Connections" -Force | Out-Null
  New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Network Connections" -Name NC_ShowSharedAccessUI -PropertyType DWORD -Value 0 -Force | Out-Null

  # Disable External Storage
  New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices' -Force | Out-Null
  New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices' -Name Deny_All -PropertyType DWORD -Value 1 -Force | Out-Null

  # USBSTOR service disabled
  New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' -Force | Out-Null
  Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' -Name Start -Type DWord -Value 4

  # Edge policies
  New-Item -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Force | Out-Null
  New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageAllowedBackgroundTypes -PropertyType DWORD -Value 3 -Force | Out-Null
  New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageContentEnabled         -PropertyType DWORD -Value 0 -Force | Out-Null
  New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageHideDefaultTopSites    -PropertyType DWORD -Value 1 -Force | Out-Null
  New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageQuickLinksEnabled      -PropertyType DWORD -Value 0 -Force | Out-Null
  New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name PrimaryPasswordSetting           -PropertyType DWORD -Value 1 -Force | Out-Null

  # Chrome policies
  New-Item -Path 'HKLM:\Software\Policies\Google\Chrome' -Force | Out-Null
  New-ItemProperty -Path 'HKLM:\Software\Policies\Google\Chrome' -Name NewTabPageLocation -PropertyType String -Value "www.google.com" -Force | Out-Null

  Write-Log "HKLM policy baselines applied."

  # --- Wallpaper file staging (and optional CSP values) ---
  $WallpaperFolder = "C:\Wallpaper"
  New-Item -ItemType Directory -Force -Path $WallpaperFolder | Out-Null
  Copy-Item -Path $jpg -Destination (Join-Path $WallpaperFolder 'MAQSoftware.jpg') -Force
  Write-Log "Wallpaper staged to C:\Wallpaper\MAQSoftware.jpg"

  # Optional CSP keys (may have limitations on Pro)
  $CSPKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
  New-Item -Path $CSPKey -Force | Out-Null
  $wall = "C:\Wallpaper\MAQSoftware.jpg"
  New-ItemProperty -Path $CSPKey -Name DesktopImageStatus    -Value 1 -PropertyType DWORD  -Force | Out-Null
  New-ItemProperty -Path $CSPKey -Name LockScreenImageStatus -Value 1 -PropertyType DWORD  -Force | Out-Null
  New-ItemProperty -Path $CSPKey -Name DesktopImagePath      -Value $wall -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $CSPKey -Name LockScreenImagePath   -Value $wall -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $CSPKey -Name DesktopImageUrl       -Value $wall -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $CSPKey -Name LockScreenImageUrl    -Value $wall -PropertyType String -Force | Out-Null
  Write-Log "Personalization CSP values written."

  # --- Schedule next phases ---
  $FirstLogonTask  = "OSDCloud-FirstLogon"
  $DeferredTask    = "OSDCloud-DeferredSystem"

  $FirstLogonScript = Join-Path $Scripts 'firstlogon.ps1'
  $DeferredScript   = Join-Path $Scripts 'deferred-system.ps1'

  Write-Log "Creating scheduled task: $FirstLogonTask (ONLOGON, user context)"
  schtasks /Create /TN $FirstLogonTask /SC ONLOGON /RL HIGHEST /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$FirstLogonScript`"" /F | Out-Null

  Write-Log "Creating scheduled task: $DeferredTask (ONSTART, SYSTEM context)"
  schtasks /Create /TN $DeferredTask /SC ONSTART /RU "SYSTEM" /RL HIGHEST /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$DeferredScript`"" /F | Out-Null

  New-Item -ItemType File -Path (Join-Path $Flags '.Setup.done') -Force | Out-Null
  Write-Log "=== Setup.ps1 completed successfully ==="
  exit 0
}
catch {
  Write-Log "ERROR: $($_.Exception.Message)"
  Write-Log "STACK: $($_.ScriptStackTrace)"
  exit 1
}
finally {
  Stop-Transcript | Out-Null
}