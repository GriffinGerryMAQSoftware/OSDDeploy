# oobe_new.ps1 - Combined SetupComplete and OOBE tasks script
# This script combines the functionality of both SetupComplete.ps1 and oobe.ps1
# It can be called from SetupComplete.cmd or run independently during OOBE

param(
    [string]$Phase = "Auto"  # "SetupComplete", "OOBE", or "Auto" (detect based on context)
)

$ErrorActionPreference = 'Stop'

# -------- Determine execution phase --------
$IsSetupComplete = $false
$IsOOBE = $false

if ($Phase -eq "SetupComplete") {
    $IsSetupComplete = $true
}
elseif ($Phase -eq "OOBE") {
    $IsOOBE = $true
}
else {
    
    $CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()

    if ($Phase -eq "SetupComplete") { $IsSetupComplete = $true }
    elseif ($Phase -eq "OOBE") { $IsOOBE = $true }
    else {
        if ($CurrentUser.IsSystem -and (Test-Path (Join-Path $PSScriptRoot 'Assets\MAQAuditor.msi'))) {
            $IsSetupComplete = $true
        }
        else {
            $IsOOBE = $true
        }
    }
    else {
        $IsOOBE = $true
    }
}

# -------- Logging Setup --------
$LogDir = 'C:\Windows\Temp'
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

if ($IsSetupComplete) {
    $LogFile = Join-Path $LogDir 'SetupComplete-OSDCloud.log'
    $TranscriptFile = Join-Path $LogDir 'SetupComplete-OSDCloud-transcript.log'
    $ScriptPhase = "SetupComplete"
}
else {
    $LogFile = Join-Path $LogDir 'OOBE-Tasks.log'
    $TranscriptFile = Join-Path $LogDir 'OOBE-Tasks-transcript.log'
    $ScriptPhase = "OOBE"
}

function Write-Log {
    param([string]$Message)
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Add-Content -Path $LogFile -Value $line
}

Start-Transcript -Path $TranscriptFile -Append | Out-Null
Write-Log "==== $ScriptPhase phase start (SYSTEM=$([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem)) ===="
Write-Log "Running as: $env:USERNAME"
Write-Log "Date/Time: $(Get-Date)"

try {
    if ($IsSetupComplete) {
        # ===============================
        # SETUPCOMPLETE PHASE
        # ===============================
        
        # ----------------------------
        # Paths setup
        # ----------------------------
        $OSDRoot = 'C:\ProgramData\OSDCloud'
        $AssetsDest = Join-Path $OSDRoot 'Assets'
        $ScriptsDest = Join-Path $OSDRoot 'Scripts'

        New-Item -Path $AssetsDest  -ItemType Directory -Force | Out-Null
        New-Item -Path $ScriptsDest -ItemType Directory -Force | Out-Null

        
        # Source is always next to this script when launched from SetupComplete.cmd
        $PayloadRoot = $PSScriptRoot
        $SrcAssets = Join-Path $PayloadRoot 'Assets'

        if (-not (Test-Path (Join-Path $SrcAssets 'MAQAuditor.msi'))) {
            Write-Log "ERROR: Required asset missing: $SrcAssets\MAQAuditor.msi"
            throw "Could not locate required assets in payload Assets folder"
        }

        Write-Log "Using payload assets from: $SrcAssets"


        # ----------------------------
        # Copy required files
        # ----------------------------
        Write-Log "Copying MAQAuditor.msi from $SrcAssets"
        Copy-Item -Path (Join-Path $SrcAssets 'MAQAuditor.msi') -Destination (Join-Path $AssetsDest 'MAQAuditor.msi') -Force

        Write-Log "Copying MAQSoftware.jpg from $SrcAssets"
        Copy-Item -Path (Join-Path $SrcAssets 'MAQSoftware.jpg') -Destination (Join-Path $AssetsDest 'MAQSoftware.jpg') -Force

        # ----------------------------
        # Sanity checks
        # ----------------------------
        if (-not (Test-Path (Join-Path $AssetsDest 'MAQAuditor.msi'))) {
            throw "MAQAuditor.msi missing after copy"
        }

        if (-not (Test-Path (Join-Path $AssetsDest 'MAQSoftware.jpg'))) {
            throw "MAQSoftware.jpg missing after copy"
        }

        # ----------------------------
        # Install Microsoft Office (ODT)
        # ----------------------------
        $OfficeSetup = Join-Path $SrcAssets 'Office\setup.exe'

        if (Test-Path $OfficeSetup) {
            Write-Log "Installing Microsoft Office via ODT"

            $OfficeDest = Join-Path $AssetsDest 'Office'
            New-Item -Path $OfficeDest -ItemType Directory -Force | Out-Null

            Write-Log "Copying Office ODT payload from $($SrcAssets)\Office to $OfficeDest"
            Copy-Item -Path (Join-Path $SrcAssets 'Office\*') -Destination $OfficeDest -Recurse -Force

            $SetupExe = Join-Path $OfficeDest 'setup.exe'
            $CfgXml = Join-Path $OfficeDest 'configuration.xml'

            if (-not (Test-Path $CfgXml)) {
                Write-Log "WARNING: configuration.xml not found at $CfgXml. Office install may fail."
            }

            # Run ODT
            Write-Log "Running: $SetupExe /configure $CfgXml"
            $p = Start-Process -FilePath $SetupExe -ArgumentList "/configure `"$CfgXml`"" -Wait -PassThru
            Write-Log "Office installer exit code: $($p.ExitCode)"
            Write-Log "Office installation completed"
        }
        else {
            Write-Log "Office installer not found. Skipping Office install."
        }

        Write-Log "SetupComplete phase completed successfully. Continuing to OOBE phase..."
    }

    # ===============================
    # OOBE PHASE (runs after SetupComplete or independently)
    # ===============================
    
    # -------- MAQAuditor MSI Installation --------
    $MsiPath = "C:\ProgramData\OSDCloud\Assets\MAQAuditor.msi"
    $MsiLog = "C:\ProgramData\OSDCloud\Assets\MAQAuditor_install.log"

    if (Test-Path $MsiPath) {
        Write-Log "Installing MAQAuditor: $MsiPath"
        $msiArgs = "/i `"$MsiPath`" /qn /norestart /L*v `"$MsiLog`""
        $proc = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru
        Write-Log "MAQAuditor msiexec exit code: $($proc.ExitCode)"
        if ($proc.ExitCode -ne 0) { 
            Write-Log "MAQAuditor install failed. ExitCode=$($proc.ExitCode). See $MsiLog"
            # Don't throw - continue with other tasks
        }
    }
    else {
        Write-Log "WARNING: MAQAuditor MSI not found at $MsiPath"
    }

    # -------- WinGet Application Installation --------
    # Microsoft docs: WinGet isn't available until first user logon triggers Store registration
    function Get-WingetPath {
        if (Get-Command winget -ErrorAction SilentlyContinue) { return "winget" }

        $candidates = Get-ChildItem "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending
        if ($candidates) { return $candidates[0].FullName }
        return $null
    }

    $wingetExe = Get-WingetPath

    $wingetApps = @(
        @{Name = "VS Code Insiders"; ID = "Microsoft.VisualStudioCodeInsiders" },
        @{Name = "Dell Command Update"; ID = "dell.commandupdate" },
        @{Name = "Adobe Acrobat Reader"; ID = "Adobe.Acrobat.Reader.64-bit" },
        @{Name = "Google Chrome"; ID = "Google.Chrome" }
    )

    if (-not $wingetExe) {
        Write-Log "WinGet not found/usable during SYSTEM context (common before first user logon). Deferring app installs."
        
        # Create a first-logon scheduled task that runs in the logged-on user's context
        $TaskScript = Join-Path $LogDir 'Install-Apps-AfterFirstLogon.ps1'
        @"
`$ErrorActionPreference = 'Continue'
`$log = '$LogDir\AfterFirstLogon-AppInstalls.log'
function wl(`$m){ Add-Content -Path `$log -Value \"[`$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] `$m\" }

wl 'Starting deferred WinGet installs...'

# Try to ensure WinGet registration (Microsoft guidance)
try {
  Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction SilentlyContinue
  wl 'Requested DesktopAppInstaller registration.'
} catch {
  wl \"Registration attempt failed: `$(`$_.Exception.Message)\"
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
        Write-Log "WinGet found at: $wingetExe"
        foreach ($app in $wingetApps) {
            Write-Log "Installing (WinGet) $($app.Name) ($($app.ID))..."
            & $wingetExe install -e --silent --disable-interactivity `
                --accept-package-agreements --accept-source-agreements `
                --id $app.ID 2>&1 | ForEach-Object { Write-Log $_ }
            Write-Log "WinGet ExitCode=$LASTEXITCODE for $($app.Name)"
            Start-Sleep -Seconds 2
        }
    }

    # -------- Registry Policies Configuration --------
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

    # USBSTOR service Start=4 (disabled)
    New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' -Force | Out-Null
    Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' -Name Start -Type DWord -Value 4

    # Edge policies
    New-Item -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageAllowedBackgroundTypes -PropertyType DWORD  -Value 3 -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageContentEnabled         -PropertyType DWORD  -Value 0 -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageHideDefaultTopSites    -PropertyType DWORD  -Value 1 -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name NewTabPageQuickLinksEnabled      -PropertyType DWORD  -Value 0 -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name PrimaryPasswordSetting           -PropertyType DWORD  -Value 1 -Force | Out-Null

    # Chrome policies
    New-Item -Path 'HKLM:\Software\Policies\Google\Chrome' -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\Policies\Google\Chrome' -Name NewTabPageLocation -PropertyType String -Value "www.google.com" -Force | Out-Null

    Write-Log "Policy registry settings applied."

    # -------- Wallpaper / Lock Screen Configuration --------
    # Note: Personalization CSP is Enterprise/Education-focused; Pro has limitations
    $WallpaperSource = "C:\ProgramData\OSDCloud\Assets\MAQSoftware.jpg"
    $WallpaperFolder = "C:\Wallpaper"
    $WallpaperPath = Join-Path $WallpaperFolder "MAQSoftware.jpg"
    $CSPKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

    New-Item -ItemType Directory -Path $WallpaperFolder -Force | Out-Null
    if (Test-Path $WallpaperSource) {
        Copy-Item -Path $WallpaperSource -Destination $WallpaperPath -Force
        Write-Log "Wallpaper copied to $WallpaperPath"
    }
    else {
        Write-Log "WARNING: Wallpaper source not found at $WallpaperSource"
    }

    New-Item -Path $CSPKey -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name DesktopImageStatus     -Value 1 -PropertyType DWORD  -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name LockScreenImageStatus  -Value 1 -PropertyType DWORD  -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name DesktopImagePath       -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name LockScreenImagePath    -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name DesktopImageUrl        -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CSPKey -Name LockScreenImageUrl     -Value $WallpaperPath -PropertyType String -Force | Out-Null
    Write-Log "PersonalizationCSP values set (note Pro limitations)"

    # -------- Printer Installation (Network-dependent, fail-soft) --------
    Write-Log "Starting printer installation section (network-dependent)."
    try {
        # Add your printer installation logic here if needed
        # This section is wrapped in try-catch to prevent blocking other OOBE tasks
        Write-Log "Printer installation section completed."
    }
    catch {
        Write-Log "WARNING: Printer installation section failed: $($_.Exception.Message)"
        # Do not throw—don't block the rest of OOBE tasks
    }

    # -------- Network Service Restart --------
    try {
        Write-Log "Restarting NlaSvc..."
        Restart-Service -Name NlaSvc -Force -ErrorAction Stop
        Write-Log "NlaSvc restarted successfully."
    }
    catch {
        Write-Log "WARNING: Could not restart NlaSvc: $($_.Exception.Message)"
    }

    Write-Log "==== $ScriptPhase tasks completed successfully ===="
}
catch {
    Write-Log "==== $ScriptPhase tasks completed with errors ===="
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log "STACK: $($_.ScriptStackTrace)"
    
    if ($IsSetupComplete) {
        exit 1
    }
    else {
        # For OOBE phase, log error but don't throw to prevent blocking
        Write-Log "Continuing despite errors in OOBE phase..."
    }
}
finally {
    Stop-Transcript | Out-Null
}

# Return appropriate exit code
if ($IsSetupComplete) {
    exit 0
}
