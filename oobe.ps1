
Start-Transcript C:\Windows\Temp\OOBE-Tasks.log


$MsiPath = "C:\ProgramData\OSDCloud\Assets\MAQAuditor.msi"
$LogPath = "C:\ProgramData\OSDCloud\Assets\MAQAuditor_install.log"

if (-not (Test-Path $MsiPath)) {
    throw "MAQAuditor MSI not found at $MsiPath"
}

$msiArgs = "/i `"$MsiPath`" /qn /norestart /L*v `"$LogPath`""
$proc = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru

if ($proc.ExitCode -ne 0) {
    throw "MAQAuditor install failed. ExitCode=$($proc.ExitCode)"
}

# --- Disable Consumer Apps ---
reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent `
/v DisableWindowsConsumerFeatures /t REG_DWORD /d 1 /f

# --- Winget ---
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "Winget not available yet"
} else {
    # VS Code Insiders
    winget install -e --silent --accept-package-agreements --accept-source-agreements --id Microsoft.VisualStudioCodeInsiders
    # Dell | Command Update
    winget install -e --silent --accept-package-agreements --accept-source-agreements --id dell.commandupdate
    # Adobe Acrobat Reader
    winget install -e --silent --accept-package-agreements --accept-source-agreements --id Adobe.Acrobat.Reader.64-bit
    # Google Chrome
    winget install -e --silent --accept-package-agreements --accept-source-agreements --id Google.Chrome
}


# --- Set Wallpaper & Lock Screen (offline, CSP) ---
$WallpaperFolder = "C:\Wallpaper"
$WallpaperPath   = Join-Path $WallpaperFolder "MAQSoftware.jpg"
$CSPKey          = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

if (!(Test-Path $WallpaperFolder)) {
    New-Item -ItemType Directory -Path $WallpaperFolder -Force | Out-Null
}

Copy-Item -Path $Wallpaper -Destination $WallpaperPath -Force

if (!(Test-Path $CSPKey)) {
    New-Item -Path $CSPKey -Force | Out-Null
}

New-ItemProperty -Path $CSPKey -Name "DesktopImageStatus"   -Value 1 -PropertyType DWORD  -Force | Out-Null
New-ItemProperty -Path $CSPKey -Name "LockScreenImageStatus" -Value 1 -PropertyType DWORD  -Force | Out-Null
New-ItemProperty -Path $CSPKey -Name "DesktopImagePath"     -Value $WallpaperPath -PropertyType String -Force | Out-Null
New-ItemProperty -Path $CSPKey -Name "LockScreenImagePath"  -Value $WallpaperPath -PropertyType String -Force | Out-Null


Stop-Transcript
