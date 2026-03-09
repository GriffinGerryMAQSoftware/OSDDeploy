
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



# --- Winget (robust in SYSTEM) ---
$wingetExe = $null

# Try PATH first
if (Get-Command winget -ErrorAction SilentlyContinue) {
    $wingetExe = "winget"
} else {
    # Try WindowsApps location (SYSTEM-readable path depends on build)
    $candidates = Get-ChildItem "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending
    if ($candidates) { $wingetExe = $candidates[0].FullName }
}

if (-not $wingetExe) {
    Write-Host "Winget not available in SYSTEM context yet; skipping app installs."
} else {
    $wingetApps = @(
        @{Name="VS Code Insiders"; ID="Microsoft.VisualStudioCodeInsiders"},
        @{Name="Dell Command Update"; ID="dell.commandupdate"},
        @{Name="Adobe Acrobat Reader"; ID="Adobe.Acrobat.Reader.64-bit"},
        @{Name="Google Chrome"; ID="Google.Chrome"}
    )

    foreach ($app in $wingetApps) {
        Write-Host "Installing $($app.Name)..."
        & $wingetExe install -e --silent --disable-interactivity `
            --accept-package-agreements --accept-source-agreements `
            --id $app.ID 2>&1 | ForEach-Object { $_ }
        Write-Host "ExitCode=$LASTEXITCODE for $($app.Name)"
        Start-Sleep -Seconds 2
    }
}
# --- CONFIGURATIONS ---

# --- Disable Consumer Apps ---
reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent `
/v DisableWindowsConsumerFeatures /t REG_DWORD /d 1 /f

# --- Disable Mobile Hotspot ---
# Define the registry keys
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Network Connections"
$registryName = "NC_ShowSharedAccessUI"

# Check if the registry keys exist
try {
    if (Test-Path $registryPath) {
        # Disable the mobile hotspot feature by setting the registry value to 0
        Set-ItemProperty -Path $registryPath -Name $registryName -Value 0
    } else {
        # Create the registry keys and disable the mobile hotspot feature
        New-Item -Path $registryPath -Force | Out-Null
        New-ItemProperty -Path $registryPath -Name $registryName -Value 0 -PropertyType DWORD | Out-Null
    }
    Write-Host "Mobile hotspot disabled successfully"
} catch {
    Write-Host "Error disabling mobile hotspot: $($_.Exception.Message)"
}

# Note: Network service restart moved to end of script to avoid disrupting network operations

# --- Disable External Storage --- 
try {
    $RemovableStorageDevicesKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices'
    if (-not (Test-Path $RemovableStorageDevicesKey)) {
        New-Item -Path $RemovableStorageDevicesKey -Force | Out-Null
    }
    Set-ItemProperty -Path $RemovableStorageDevicesKey -Name "Deny_All" -Value 1 -Type DWORD -Force | Out-Null
    
    $UsbStorKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR'
    if (-not (Test-Path $UsbStorKey)) {
        New-Item -Path $UsbStorKey -Force | Out-Null
    }
    Set-ItemProperty -Path $UsbStorKey -Name "Start" -Value 4 -Type DWORD -Force | Out-Null
    Write-Host "External storage disabled successfully"
} catch {
    Write-Host "Error disabling external storage: $($_.Exception.Message)"
}

# --- Set Edge and Chrome New Tab Policies ---
# Define registry keys for Edge and Chrome policies
$EdgeRegKey = 'HKLM:\Software\Policies\Microsoft\Edge'
$ChromeRegKey = 'HKLM:\Software\Policies\Google\Chrome'

# Function to create a registry key if it does not exist
function Ensure-RegistryKeyExists {
param (
[string]$RegKeyPath
)
if (-Not (Test-Path $RegKeyPath)) {
New-Item -Path $RegKeyPath -Force | Out-Null
Write-Host "Created registry key: $RegKeyPath"
} else {
Write-Host "Registry key already exists: $RegKeyPath"
}
}

# Function to set a registry property
function Set-RegistryProperty {
param (
[string]$RegKeyPath,
[string]$PropertyName,
[string]$PropertyValue,
[string]$PropertyType
)
New-ItemProperty -Path $RegKeyPath -Name $PropertyName -Value $PropertyValue -PropertyType $PropertyType -Force | Out-Null
Write-Host "Set $PropertyName to $PropertyValue in $RegKeyPath"
}

# Ensure Edge registry key exists
Ensure-RegistryKeyExists -RegKeyPath $EdgeRegKey

# Set Edge policies
# Set Edge policies

# Allow background types on the new tab page (3 = Image and Video)
Set-RegistryProperty -RegKeyPath $EdgeRegKey -PropertyName "NewTabPageAllowedBackgroundTypes" -PropertyValue "3" -PropertyType "DWORD"

# Enable or disable content on the new tab page (0 = Disabled)
Set-RegistryProperty -RegKeyPath $EdgeRegKey -PropertyName "NewTabPageContentEnabled" -PropertyValue "0" -PropertyType "DWORD"

# Hide default top sites on the new tab page (1 = Hide)
Set-RegistryProperty -RegKeyPath $EdgeRegKey -PropertyName "NewTabPageHideDefaultTopSites" -PropertyValue "1" -PropertyType "DWORD"

# Enable or disable quick links on the new tab page (0 = Disabled)
Set-RegistryProperty -RegKeyPath $EdgeRegKey -PropertyName "NewTabPageQuickLinksEnabled" -PropertyValue "0" -PropertyType "DWORD"

# Require password when filling passwords from password manager (2 = With device password)
Set-RegistryProperty -RegKeyPath $EdgeRegKey -PropertyName "PrimaryPasswordSetting" -PropertyValue "1" -PropertyType "DWORD"


# Ensure Chrome registry key exists
Ensure-RegistryKeyExists -RegKeyPath $ChromeRegKey

# Set Chrome policies

# Set the URL for the new tab page
Set-RegistryProperty -RegKeyPath $ChromeRegKey -PropertyName "NewTabPageLocation" -PropertyValue "www.google.com" -PropertyType "STRING"


# --- Set Wallpaper & Lock Screen (offline, CSP) ---
$WallpaperSource = "C:\ProgramData\OSDCloud\Assets\MAQSoftware.jpg"  # Source wallpaper file
$WallpaperFolder = "C:\Wallpaper"
$WallpaperPath   = Join-Path $WallpaperFolder "MAQSoftware.jpg"
$CSPKey          = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

try {
    if (!(Test-Path $WallpaperFolder)) {
        New-Item -ItemType Directory -Path $WallpaperFolder -Force | Out-Null
    }

    # Check if source wallpaper exists before copying
    if (Test-Path $WallpaperSource) {
        Copy-Item -Path $WallpaperSource -Destination $WallpaperPath -Force
    } else {
        Write-Host "Warning: Wallpaper source not found at $WallpaperSource"
    }
} catch {
    Write-Host "Error setting wallpaper: $($_.Exception.Message)"
}

if (!(Test-Path $CSPKey)) {
    New-Item -Path $CSPKey -Force | Out-Null
}

New-ItemProperty -Path $CSPKey -Name "DesktopImageStatus"   -Value 1 -PropertyType DWORD  -Force | Out-Null
New-ItemProperty -Path $CSPKey -Name "LockScreenImageStatus" -Value 1 -PropertyType DWORD  -Force | Out-Null
New-ItemProperty -Path $CSPKey -Name "DesktopImagePath"     -Value $WallpaperPath -PropertyType String -Force | Out-Null
New-ItemProperty -Path $CSPKey -Name "LockScreenImagePath"  -Value $WallpaperPath -PropertyType String -Force | Out-Null
New-ItemProperty -Path $CSPKey -Name "DesktopImageUrl"     -Value $WallpaperPath -PropertyType String -Force | Out-Null
New-ItemProperty -Path $CSPKey -Name "LockScreenImageUrl"  -Value $WallpaperPath -PropertyType String -Force | Out-Null

# --- Download and Install Printer Drivers ---
# Set absolute download path for SYSTEM context
$script:PrinterDownloadRoot = "C:\Windows\Temp\PrinterDrivers"

$printersTable = @{
    0 = @{
        ip = "192.168.2.7"
        portName = "IP_192.168.2.7"
        printerName = "MAQ Software - Redmond Printer"
        driverName = "Brother MFC-L6700DW series"
        driverURL = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/03_MFC-L6xxx_Monochrome_for_Redmond_IT.zip"
        infFile = "BRPRM15A.INF"
    }
    1 = @{
        ip = "192.168.2.3"
        portName = "IP_192.168.2.3"
        printerName = "MAQ Software - Redmond Color Printer"
        driverName = "Brother HL-L8360CDW series"
        driverURL = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/04_HL-L8xxx_ColorLaser_for_Redmond_Color.zip"
        infFile = "BROCH16A.INF"
    }
    2 = @{
        ip = "192.168.2.4"
        portName = "IP_192.168.2.4"
        printerName = "MAQ Software - Redmond HR Printer"
        driverName = "Brother MFC-L8900CDW series"
        driverURL = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/02_MFC-L8xxx_ColorSeries_for_HR.zip"
        infFile = "BRPRC16A.INF"
    }
    3 = @{
        ip = "192.168.2.6"
        portName = "IP_192.168.2.6"
        printerName = "MAQ Software - Redmond IT Printer"
        driverName = "Brother MFC-L6700DW series"
        driverURL = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/03_MFC-L6xxx_Monochrome_for_Redmond_IT.zip"
        infFile = "BRPRM15A.INF"
    }
    4 = @{
        ip = "192.168.2.5"
        portName = "IP_192.168.2.5"
        printerName = "MAQ Software - Plano Printer"
        driverName = "Brother MFC-L3780CDW series"
        driverURL = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/01_MFC-L3xxx_ColorSeries_for_Plano.zip"
        infFile = "BRPRC20A.INF"
    }
}

########################################################################################################
# Functions section
########################################################################################################

# Function to install a specific printer by index
function Install-PrinterByIndex {
    param ($printerIndex)
    
    $printerInfor = $script:printersTable[$printerIndex]
    
    # Remove existing printer instance
    RemovePrinterIfExists $printerInfor['printerName']
    
    # Remove old port
    DeleteOldPort $printerInfor['portName']
    
    # Setup new port
    SetupNewPort $printerInfor['ip'] $printerInfor['portName']
    
    # Setup Printer Driver
    SetupPrinterDriver $printerInfor['driverName'] $printerInfor['driverURL'] $printerInfor['infFile']
    
    # Check if an alternative driver name was set during installation
    $finalDriverName = if ($script:actualDriverName) { $script:actualDriverName } else { $printerInfor['driverName'] }
    
    # Setup Printer
    SetupPrinter $printerInfor['printerName'] $finalDriverName $printerInfor['portName']
    
    # Auto-delete downloaded files
    if ($script:downloadedDriverPath -and (Test-Path $script:downloadedDriverPath)) {
        try {
            Remove-Item -Path $script:downloadedDriverPath -Recurse -Force -ErrorAction Stop
        } catch {
            # Silently continue on cleanup error
        }
        $script:downloadedDriverPath = $null
    }
}

# Function to expand compressed driver files
function Expand-DriverFiles {
    param ($sourcePath)
    
    $compressedFiles = Get-ChildItem $sourcePath -Recurse -File | Where-Object { $_.Extension -match '^\.[a-z]+_$' }
    
    if ($compressedFiles) {
        $expandedCount = 0
        $extensionMap = @{
            '.dl_' = '.dll'
            '.ex_' = '.exe'
            '.in_' = '.ini'
            '.sy_' = '.sys'
            '.da_' = '.dat'
            '.ch_' = '.chm'
            '.ds_' = '.ds'
        }
        
        foreach ($compFile in $compressedFiles) {
            try {
                $targetExt = $extensionMap[$compFile.Extension.ToLower()]
                if (-not $targetExt) {
                    $targetExt = $compFile.Extension.Substring(0, $compFile.Extension.Length - 1) + $compFile.Extension[-1]
                }
                
                $targetFile = Join-Path $compFile.DirectoryName ($compFile.BaseName + $targetExt)
                
                # Skip if already expanded, but remove the compressed version
                if (Test-Path $targetFile) {
                    try {
                        Remove-Item $compFile.FullName -Force -ErrorAction SilentlyContinue
                    } catch {}
                    continue
                }
                
                $expandResult = & expand.exe $compFile.FullName $targetFile 2>&1
                if ($LASTEXITCODE -eq 0 -and (Test-Path $targetFile)) {
                    $expandedCount++
                    # Remove the compressed source file after successful expansion
                    try {
                        Remove-Item $compFile.FullName -Force -ErrorAction SilentlyContinue
                    } catch {}
                } else {
                    # Continue on error
                }
            } catch {
                # Continue on error
            }
        }
    }
}

# Function to delete old port
function DeleteOldPort {
    param ($portName)
    $portExists = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    if ($portExists) {
        try {
            Remove-PrinterPort -Name $portName -ErrorAction Stop
        } catch {
            # Silently continue if port removal fails
        }
    }
}

# Function to setup new port
function SetupNewPort {
    param ($ip, $portName)
    
    # Check if port already exists
    $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    
    if ($existingPort) {
        # Optionally verify the IP matches
        try {
            $cimPort = Get-CimInstance -ClassName Win32_TCPIPPrinterPort | Where-Object { $_.Name -eq $portName }
        } catch {
            # If we can't verify IP, just proceed
        }
        return
    }
    
    # Port doesn't exist, create it
    # Try standard method first
    try {
        Add-PrinterPort -Name $portName -PrinterHostAddress $ip -PortNumber 9100 -ErrorAction Stop
    } catch {
        # Fallback to CIM method
        try {
            $portProperties = @{
                Name = $portName
                Protocol = [uint32]1
                HostAddress = $ip
                PortNumber = [uint32]9100
                SNMPEnabled = $false
            }
            New-CimInstance -ClassName Win32_TCPIPPrinterPort -Property $portProperties -ErrorAction Stop | Out-Null
        } catch {
            throw
        }
    }
}

# Function to setup printer driver
function SetupPrinterDriver {
    param ($driverName, $driverUrl, $infFile)

    # Check if driver already exists
    $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
    if ($existingDriver) {
        return
    }
    
    # Install driver
    try {
        # Create download directory using absolute path
        $safePrinterName = $driverName -replace '[^a-zA-Z0-9-]', '_'
        $downloadPath = Join-Path $script:PrinterDownloadRoot $safePrinterName
        New-Item -ItemType Directory -Path $downloadPath -Force | Out-Null
        
        $zipFile = Join-Path $downloadPath ($driverUrl | Split-Path -Leaf)
        $extractPath = Join-Path $downloadPath "extracted"
        
        # Track download path for cleanup
        $script:downloadedDriverPath = $downloadPath

        # Download driver
        Invoke-WebRequest -Uri $driverUrl -OutFile $zipFile -ErrorAction Stop
        
        # Extract driver
        Expand-Archive -Path $zipFile -DestinationPath $extractPath -Force
        
        # Find INF file - search recursively in case it's in a subfolder
        $infPath = Join-Path $extractPath $infFile
        if (-not (Test-Path $infPath)) {
            # Try searching recursively for the specific INF file
            $alternativeInf = Get-ChildItem -Path $extractPath -Filter $infFile -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($alternativeInf) {
                $infPath = $alternativeInf.FullName
            } else {
                # Try any INF file as last resort
                $alternativeInf = Get-ChildItem "$extractPath\*.INF" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($alternativeInf) {
                    $infPath = $alternativeInf.FullName
                } else {
                    throw "No INF files found in driver package"
                }
            }
        }
        
        # Expand compressed driver files
        Expand-DriverFiles $extractPath

        # Install driver package using pnputil (add to driver store)
        $pnpResult = & pnputil.exe /add-driver $infPath /subdirs 2>&1
        
        # Parse INF file to find actual driver names
        $infContent = Get-Content $infPath -Raw
        $driverNamesInInf = @()
        
        # Look for DriverName sections (Brother INF format)
        # Try multiple section patterns: [DriverName.NTamd64], [DriverName.NTx86], [Models], etc.
        $sectionPatterns = @('DriverName\.NTamd64', 'DriverName\.NTx86', 'DriverName', '.*Models.*')
        
        foreach ($pattern in $sectionPatterns) {
            if ($infContent -match "\[$pattern\]") {
                $sections = $infContent -split '\[' | Where-Object { $_ -match "^$pattern" }
                if ($sections) {
                    $lines = $sections[0] -split "`n"
                    foreach ($line in $lines) {
                        # Match lines with quoted printer names: "Brother MFC-L3780CDW series" = ...
                        # Allow optional whitespace at start of line
                        if ($line -match '^\s*"([^"]+)"') {
                            $foundDriverName = $matches[1]
                            if ($foundDriverName -and $driverNamesInInf -notcontains $foundDriverName) {
                                $driverNamesInInf += $foundDriverName
                            }
                        }
                    }
                }
            }
        }
        
        # Extract model number from driver name to search dynamically (e.g., "L6700DW", "L3780CDW")
        $matchedDriverName = $null
        # Match pattern like "MFC-L6700DW" or "HL-L8360CDW" - extract the L#### part
        if ($driverName -match '((?:MFC|HL|DCP)-L[A-Z0-9]+)') {
            $modelNumber = $matches[1]
            $matchedDriverName = $driverNamesInInf | Where-Object { $_ -like "*$modelNumber*" } | Select-Object -First 1
        }
        
        # If no match by model number, try exact driver name
        if (-not $matchedDriverName) {
            $matchedDriverName = $driverNamesInInf | Where-Object { $_ -eq $driverName } | Select-Object -First 1
        }
        
        # Register driver with print system using Add-PrinterDriver
        $registered = $false
        
        # Try the matched driver first by name (after pnputil added it to store)
        if ($matchedDriverName) {
            try {
                Add-PrinterDriver -Name $matchedDriverName -ErrorAction Stop
                $script:actualDriverName = $matchedDriverName
                $registered = $true
            } catch {
                # Continue to try other drivers
            }
        }
        
        # If matched driver failed, try all drivers from INF by name
        if (-not $registered) {
            foreach ($driverNameToTry in $driverNamesInInf) {
                # Skip if we already tried this one
                if ($matchedDriverName -and $driverNameToTry -eq $matchedDriverName) {
                    continue
                }
                
                try {
                    Add-PrinterDriver -Name $driverNameToTry -ErrorAction Stop
                    $script:actualDriverName = $driverNameToTry
                    $registered = $true
                    break
                } catch {
                    # Try next one
                }
            }
        }
        
        # Wait for driver to be registered
        Start-Sleep -Seconds 2
        
        # Verify installation - check both original name and script variable
        $driverToCheck = if ($script:actualDriverName) { $script:actualDriverName } else { $driverName }
        $installedDriver = Get-PrinterDriver -Name $driverToCheck -ErrorAction SilentlyContinue
        
        if (-not $installedDriver) {
            # Check for similar Brother drivers
            $brotherDrivers = Get-PrinterDriver | Where-Object { $_.Name -like "*Brother*" -or $_.Name -like "*MFC*" }
            
            if ($brotherDrivers) {
                # Try to find best match for L3780
                $bestMatch = $brotherDrivers | Where-Object { $_.Name -like "*L3780*" -or $_.Name -like "*MFC-L3780*" } | Select-Object -First 1
                if (-not $bestMatch) {
                    $bestMatch = $brotherDrivers | Select-Object -First 1
                }
                
                $script:actualDriverName = $bestMatch.Name
            } else {
                throw "Driver installation failed - no compatible driver found"
            }
        }
    }
    catch {
        throw
    }
}

# Function to setup printer
function SetupPrinter {
    param ($printerName, $driverName, $portName)
    try {
        Add-Printer -Name $printerName -DriverName $driverName -PortName $portName -ErrorAction Stop
    }
    catch {
        throw
    }
}

# Function to remove printer if it exists
function RemovePrinterIfExists {
    param ($printerName)
    if (Get-Printer -Name $printerName -ErrorAction SilentlyContinue) {
        try {
            Remove-Printer -Name $printerName -ErrorAction Stop
            
            # Verify removal with retry
            $retries = 0
            $maxRetries = 5
            while ($retries -lt $maxRetries) {
                Start-Sleep -Seconds 1
                if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                    break
                }
                $retries++
            }
        }
        catch {
            # Silently continue if removal fails
        }
    }
}

try {
    # Define printers to install (array indices: 0=Redmond Printer, 1=Redmond Color Printer, 4=Plano Printer)
    $printersToInstall = @(0, 1, 4)  # Redmond Printer, Redmond Color Printer, Plano Printer
    
    foreach ($printerIndex in $printersToInstall) {
        try {
            Install-PrinterByIndex $printerIndex
        }
        catch {
            Write-Host "Error installing printer index $printerIndex : $($_.Exception.Message)"
            # Continue with next printer instead of stopping completely
            continue
        }
    }
}
catch {
    Write-Host "Error in printer installation section: $($_.Exception.Message)"
    # Keep downloaded files for troubleshooting
    if ($script:downloadedDriverPath -and (Test-Path $script:downloadedDriverPath)) {
        Write-Host "Printer driver files kept at: $script:downloadedDriverPath"
    }
}

# --- Final Network Service Restart ---
# Restart network service at the end to apply all network-related policy changes
try {
    Write-Host "Restarting Network Location Awareness service to apply network policies..."
    Restart-Service -Name NlaSvc -Force -ErrorAction Stop
    
    # Wait for service to stabilize
    Start-Sleep -Seconds 5
    
    # Verify service is running
    $nlsStatus = Get-Service -Name NlaSvc -ErrorAction SilentlyContinue
    if ($nlsStatus -and $nlsStatus.Status -eq 'Running') {
        Write-Host "Network Location Awareness service restarted successfully"
    } else {
        Write-Host "Warning: Network Location Awareness service may not be running properly"
    }
} catch {
    Write-Host "Warning: Could not restart Network Location Awareness service: $($_.Exception.Message)"
}

Write-Host "OOBE configuration script completed"
Stop-Transcript
