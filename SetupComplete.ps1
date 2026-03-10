# SetupComplete.ps1

$ErrorActionPreference = 'Stop'

# ----------------------------
# Paths / Logging
# ----------------------------
$LogPath     = 'C:\Windows\Temp\SetupComplete-OSDCloud.log'
$Transcript  = 'C:\Windows\Temp\SetupComplete-OSDCloud-transcript.log'

$OSDRoot     = 'C:\ProgramData\OSDCloud'
$AssetsDest  = Join-Path $OSDRoot 'Assets'
$ScriptsDest = Join-Path $OSDRoot 'Scripts'

New-Item -Path $AssetsDest  -ItemType Directory -Force | Out-Null
New-Item -Path $ScriptsDest -ItemType Directory -Force | Out-Null

function Write-Log {
    param([string]$Message)
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Add-Content -Path $LogPath -Value $line
}

Write-Log "==== SetupComplete start ===="
Write-Log "Running as: $env:USERNAME"
Write-Log "Date/Time: $(Get-Date)"

Start-Transcript -Path $Transcript -Append | Out-Null

try {
    # ----------------------------
    # Prefer local staged config first (matches your batch logic)
    # ----------------------------
    $SrcAssets  = 'C:\OSDCloud\Config\Assets'
    $SrcScripts = 'C:\OSDCloud\Config\Scripts'

    if (Test-Path (Join-Path $SrcAssets 'MAQAuditor.msi')) {
        Write-Log "Found local staged assets at $SrcAssets"
    }
    else {
        Write-Log "Local staged assets not found. Scanning drives for OSDCloud media..."

        $SrcAssets  = $null
        $SrcScripts = $null

        # Match your drive scan behavior (C..Z except X is fine; we’ll scan actual filesystem drives)
        $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -match '^[A-Z]$' } | Sort-Object Name
        foreach ($d in $drives) {
            $candidateAssets  = "$($d.Name):\OSDCloud\Config\Assets"
            $candidateScripts = "$($d.Name):\OSDCloud\Config\Scripts"

            if (Test-Path (Join-Path $candidateAssets 'MAQAuditor.msi')) {
                $SrcAssets  = $candidateAssets
                $SrcScripts = $candidateScripts
                Write-Log "Found OSDCloud media at $($d.Name):\OSDCloud\Config"
                break
            }
        }
    }

    if (-not $SrcAssets) {
        Write-Log "ERROR: Could not locate MAQAuditor.msi on local disk or any drive."
        Write-Log "==== SetupComplete end (FAIL) ===="
        exit 1
    }

    # ----------------------------
    # Copy MAQAuditor MSI + wallpaper
    # ----------------------------
    Write-Log "Copying MAQAuditor.msi from $SrcAssets"
    Copy-Item -Path (Join-Path $SrcAssets 'MAQAuditor.msi') -Destination (Join-Path $AssetsDest 'MAQAuditor.msi') -Force

    Write-Log "Copying MAQSoftware.jpg from $SrcAssets"
    Copy-Item -Path (Join-Path $SrcAssets 'MAQSoftware.jpg') -Destination (Join-Path $AssetsDest 'MAQSoftware.jpg') -Force

    # Copy OOBE script
    Write-Log "Copying oobe.ps1 from $SrcScripts"
    Copy-Item -Path (Join-Path $SrcScripts 'oobe.ps1') -Destination (Join-Path $ScriptsDest 'oobe.ps1') -Force

    # ----------------------------
    # Sanity checks (same as batch)
    # ----------------------------
    if (-not (Test-Path (Join-Path $AssetsDest 'MAQAuditor.msi'))) {
        Write-Log "ERROR: MAQAuditor.msi missing after copy"
        Write-Log "==== SetupComplete end (FAIL) ===="
        exit 1
    }

    if (-not (Test-Path (Join-Path $AssetsDest 'MAQSoftware.jpg'))) {
        Write-Log "ERROR: MAQSoftware.jpg missing after copy"
        Write-Log "==== SetupComplete end (FAIL) ===="
        exit 1
    }

    if (-not (Test-Path (Join-Path $ScriptsDest 'oobe.ps1'))) {
        Write-Log "ERROR: oobe.ps1 missing after copy"
        Write-Log "==== SetupComplete end (FAIL) ===="
        exit 1
    }

    # ----------------------------
    # Install Microsoft Office (ODT) - same logic as batch
    # ----------------------------
    $OfficeSetup = Join-Path $SrcAssets 'Office\setup.exe'
    $OfficeCfg   = Join-Path $SrcAssets 'Office\configuration.xml'

    if (Test-Path $OfficeSetup) {
        Write-Log "Installing Microsoft Office via ODT"

        $OfficeDest = Join-Path $AssetsDest 'Office'
        New-Item -Path $OfficeDest -ItemType Directory -Force | Out-Null

        Write-Log "Copying Office ODT payload from $($SrcAssets)\Office to $OfficeDest"
        Copy-Item -Path (Join-Path $SrcAssets 'Office\*') -Destination $OfficeDest -Recurse -Force

        $SetupExe = Join-Path $OfficeDest 'setup.exe'
        $CfgXml   = Join-Path $OfficeDest 'configuration.xml'

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

    # ----------------------------
    # Execute OOBE tasks (your existing behavior)
    # ----------------------------
    $OobePs1 = Join-Path $ScriptsDest 'oobe.ps1'
    Write-Log "Executing oobe.ps1 from $OobePs1"
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File $OobePs1 2>&1 | ForEach-Object { Write-Log $_ }

    Write-Log "==== SetupComplete end ===="
    exit 0
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log "STACK: $($_.ScriptStackTrace)"
    Write-Log "==== SetupComplete end (FAIL) ===="
    exit 1
}
finally {
    Stop-Transcript | Out-Null
}