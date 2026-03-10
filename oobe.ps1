# oobe.ps1 (runs in SYSTEM context when invoked from SetupComplete)
$ErrorActionPreference = 'Stop'

# -------- Logging --------
$LogDir = 'C:\Windows\Temp'
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

$LogFile       = Join-Path $LogDir 'OOBE-Tasks.log'
$TranscriptFile = Join-Path $LogDir 'OOBE-Tasks-transcript.log'

function Write-Log {
    param([string]$Message)
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Add-Content -Path $LogFile -Value $line
}

Start-Transcript -Path $TranscriptFile -Append | Out-Null
Write-Log "==== OOBE tasks start (SYSTEM=$([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem)) ===="

try {
    # -------- MAQAuditor MSI --------
    $MsiPath = "C:\ProgramData\OSDCloud\Assets\MAQAuditor.msi"
    $MsiLog  = "C:\ProgramData\OSDCloud\Assets\MAQAuditor_install.log"

    if (-not (Test-Path $MsiPath)) { throw "MAQAuditor MSI not found at $MsiPath" }

    Write-Log "Installing MAQAuditor: $MsiPath"
    $msiArgs = "/i `"$MsiPath`" /qn /norestart /L*v `"$MsiLog`""
    $proc = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru
    Write-Log "MAQAuditor msiexec exit code: $($proc.ExitCode)"
    if ($proc.ExitCode -ne 0) { throw "MAQAuditor install failed. ExitCode=$($proc.ExitCode). See $MsiLog" }

    # -------- WinGet (Reality check + deferral) --------
    # Microsoft docs: WinGet isn't available until first user logon triggers Store registration. [3](https://learn.microsoft.com/en-us/windows/package-manager/winget/)
    # We'll try to locate winget.exe; if not usable, schedule installs at first logon.

    function Get-WingetPath {
        if (Get-Command winget -ErrorAction SilentlyContinue) { return "winget" }

        $candidates = Get-ChildItem "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending
        if ($candidates) { return $candidates[0].FullName }
        return $null
    }

    $wingetExe = Get-WingetPath

    $wingetApps = @(
        @{Name="VS Code Insiders";       ID="Microsoft.VisualStudioCodeInsiders"},
        @{Name="Dell Command Update";    ID="dell.commandupdate"},
        @{Name="Adobe Acrobat Reader";   ID="Adobe.Acrobat.Reader.64-bit"},
        @{Name="Google Chrome";          ID="Google.Chrome"}
    )

    if (-not $wingetExe) {
        Write-Log "WinGet not found/usable during SetupComplete (common before first user logon). Deferring app installs. [3](https://learn.microsoft.com/en-us/windows/package-manager/winget/)"
        $defer = $true
    } else {
        Write-Log "WinGet path resolved: $wingetExe"
        $defer = $false
    }

    if ($defer) {
        # Create a first-logon scheduled task that runs in the logged-on user's context
        $TaskScript = Join-Path $LogDir 'Install-Apps-AfterFirstLogon.ps1'
        @"
`$ErrorActionPreference = 'Continue'
`$log = '$LogDir\AfterFirstLogon-AppInstalls.log'
function wl(`$m){ Add-Content -Path `$log -Value \"[`$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] `$m\" }

wl 'Starting deferred WinGet installs...'

# Try to ensure WinGet registration (Microsoft guidance) [3](https://learn.microsoft.com/en-us/windows/package-manager/winget/)
try {
  Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction SilentlyContinue
  wl 'Requested DesktopAppInstaller registration.'
} catch {
  wl \"Registration attempt failed: `$($_.Exception.Message)\"
}

# Resolve winget again
`$winget = (Get-Command winget -ErrorAction SilentlyContinue).Source
if (-not `$winget) {
  `$c = Get-ChildItem \"C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe\" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (`$c) { `$winget = `$c.FullName }
}

if (-not `$winget) {
  wl 'WinGet still not available; exiting.'
  exit 0
}

wl \"WinGet resolved: `$winget\"

`$apps = @(
  @{Name='VS Code Insiders'; ID='Microsoft.VisualStudioCodeInsiders'},
  @{Name='Dell Command Update'; ID='dell.commandupdate'},
  @{Name='Adobe Acrobat Reader'; ID='Adobe.Acrobat.Reader.64-bit'},
  @{Name='Google Chrome'; ID='Google.Chrome'}
)

foreach (`$app in `$apps) {
  wl \"Installing `$(`$app.Name) (`$(`$app.ID))...\"
  & `$winget install -e --silent --disable-interactivity --accept-package-agreements --accept-source-agreements --id `$app.ID *>> `$log
  wl \"ExitCode=`$LASTEXITCODE for `$(`$app.Name)\"
}
wl 'Deferred installs complete.'
"@ | Set-Content -Path $TaskScript -Encoding UTF8

        $TaskName = "MAQ-DeferredAppInstalls"
        Write-Log "Creating scheduled task '$TaskName' to run at first logon: $TaskScript"

        # Runs when any user logs on; executes in that user's context
        schtasks /Create /TN $TaskName /SC ONLOGON /RL HIGHEST /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$TaskScript`"" /F | Out-Null
        Write-Log "Scheduled task created."
    }
    else {
        foreach ($app in $wingetApps) {
            Write-Log "Installing (WinGet) $($app.Name) ($($app.ID))..."
            & $wingetExe install -e --silent --disable-interactivity `
                --accept-package-agreements --accept-source-agreements `
                --id $app.ID 2>&1 | ForEach-Object { Write-Log $_ }
            Write-Log "WinGet ExitCode=$LASTEXITCODE for $($app.Name)"
            Start-Sleep -Seconds 2
        }
    }

    # -------- Policies / Registry --------
    Write-Log "Applying policy registry settings..."

    # Disable Consumer Apps
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Force | Out-Null
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" `
        -Name DisableWindowsConsumerFeatures -PropertyType DWORD -Value 1 -Force | Out-Null

    # Disable Mobile Hotspot UI
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Network Connections" -Force | Out-Null
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Network Connections" `
        -Name NC_ShowSharedAccessUI -PropertyType DWORD -Value 0 -Force | Out-Null

    # Disable External Storage
    New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices' -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices' `
        -Name Deny_All -PropertyType DWORD -Value 1 -Force | Out-Null

    # USBSTOR service Start=4
    New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' -Force | Out-Null
    Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' -Name Start -Type DWord -Value 4

    # Edge + Chrome policies
    New-Item -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageAllowedBackgroundTypes -PropertyType DWORD  -Value 3 -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageContentEnabled         -PropertyType DWORD  -Value 0 -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageHideDefaultTopSites    -PropertyType DWORD  -Value 1 -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageQuickLinksEnabled      -PropertyType DWORD  -Value 0 -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name PrimaryPasswordSetting           -PropertyType DWORD  -Value 1 -Force | Out-Null

    New-Item -Path 'HKLM:\Software\Policies\Google\Chrome' -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Google\Chrome' -Name NewTabPageLocation -PropertyType String -Value "www.google.com" -Force | Out-Null

    Write-Log "Policy registry settings applied."

    # -------- Wallpaper / Lock Screen via PersonalizationCSP --------
    # Note: Personalization CSP is Enterprise/Education-focused; Pro has limitations. [2](https://learn.microsoft.com/en-us/windows/client-management/mdm/personalization-csp)
    $WallpaperSource = "C:\ProgramData\OSDCloud\Assets\MAQSoftware.jpg"
    $WallpaperFolder = "C:\Wallpaper"
    $WallpaperPath   = Join-Path $WallpaperFolder "MAQSoftware.jpg"
    $CSPKey          = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

    New-Item -ItemType Directory -Path $WallpaperFolder -Force | Out-Null
    if (Test-Path $WallpaperSource) {
        Copy-Item -Path $WallpaperSource -Destination $WallpaperPath -Force
        Write-Log "Wallpaper copied to $WallpaperPath"
    } else {
        Write-Log "WARNING: Wallpaper source not found at $WallpaperSource"
    }

    New-Item -Path $CSPKey -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name DesktopImageStatus     -Value 1 -PropertyType DWORD  -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name LockScreenImageStatus  -Value 1 -PropertyType DWORD  -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name DesktopImagePath       -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name LockScreenImagePath    -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name DesktopImageUrl        -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name LockScreenImageUrl     -Value $WallpaperPath -PropertyType String -Force | Out-Null
    Write-Log "PersonalizationCSP values set (note Pro limitations). [2](https://learn.microsoft.com/en-us/windows/client-management/mdm/personalization-csp)"

    # -------- Printer installs (fail-soft) --------
    # Your existing printer function block is fine, but it’s network-dependent and can fail during SetupComplete.
    # Keep your existing functions/table here, but wrap the whole section:
    Write-Log "Starting printer installation section (network-dependent)."
    try {
        # <<< KEEP YOUR EXISTING $printersTable + functions + foreach loop HERE >>>
        Write-Log "Printer installation section completed."
    }
    catch {
        Write-Log "WARNING: Printer installation section failed: $($_.Exception.Message)"
        # Do not throw—don’t block the rest of OOBE tasks
    }

    # -------- Optional: restart NlaSvc at very end --------
    try {
        Write-Log "Restarting NlaSvc..."
        Restart-Service -Name NlaSvc -Force -ErrorAction Stop
        Write-Log "NlaSvc restarted."
    } catch {
        Write-Log "WARNING: Could not restart NlaSvc: $($_.Exception.Message)"
    }

    Write-Log "==== OOBE tasks complete (SUCCESS) ===="
}
catch {
    Write-Log "==== OOBE tasks complete (FAIL) ===="
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log "STACK: $($_.ScriptStackTrace)"
    throw
}
finally {
    Stop-Transcript | Out-Null
}