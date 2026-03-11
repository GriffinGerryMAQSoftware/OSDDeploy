# Deferred-System.ps1 (SYSTEM)
# Purpose: silent machine installs/config after setup without blocking SetupComplete UI.

$ErrorActionPreference = 'Continue'

$Root = 'C:\ProgramData\OSDCloud'
$Assets = Join-Path $Root 'Assets'
$Logs = Join-Path $Root 'Logs'
$Flags = Join-Path $Root 'Flags'

New-Item -ItemType Directory -Force -Path $Logs, $Flags | Out-Null

$LogFile = Join-Path $Logs 'Deferred-System.log'
Start-Transcript -Path (Join-Path $Logs 'Deferred-System-transcript.log') -Append | Out-Null

function Write-Log([string]$m) {
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
    $MsiLog = Join-Path $Assets 'MAQAuditor_install.log'

    if (Test-Path $MsiPath) {
        Write-Log "Installing MAQAuditor MSI: $MsiPath"
        $args = "/i `"$MsiPath`" /qn /norestart /L*v `"$MsiLog`""

        $p = Start-Process msiexec.exe -ArgumentList $args -PassThru
        $timeoutSec = 1200

        if (Wait-Process -Id $p.Id -Timeout $timeoutSec -ErrorAction SilentlyContinue) {
            Write-Log "MAQAuditor install finished. ExitCode=$($p.ExitCode)"
        }
        else {
            Write-Log "WARNING: MAQAuditor install exceeded ${timeoutSec}s; continuing."
            try { Stop-Process -Id $p.Id -Force } catch {}
        }
    }
    else {
        Write-Log "WARNING: MAQAuditor MSI not found at $MsiPath"
    }

    # --- Download and Install Printer Drivers ---
    # Set absolute download path for SYSTEM context
    $script:PrinterDownloadRoot = "C:\Windows\Temp\PrinterDrivers"

    $printersTable = @{
        0 = @{
            ip          = "192.168.2.7"
            portName    = "IP_192.168.2.7"
            printerName = "MAQ Software - Redmond Printer"
            driverName  = "Brother MFC-L6700DW series"
            driverURL   = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/03_MFC-L6xxx_Monochrome_for_Redmond_IT.zip"
            infFile     = "BRPRM15A.INF"
        }
        1 = @{
            ip          = "192.168.2.3"
            portName    = "IP_192.168.2.3"
            printerName = "MAQ Software - Redmond Color Printer"
            driverName  = "Brother HL-L8360CDW series"
            driverURL   = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/04_HL-L8xxx_ColorLaser_for_Redmond_Color.zip"
            infFile     = "BROCH16A.INF"
        }
        2 = @{
            ip          = "192.168.2.4"
            portName    = "IP_192.168.2.4"
            printerName = "MAQ Software - Redmond HR Printer"
            driverName  = "Brother MFC-L8900CDW series"
            driverURL   = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/02_MFC-L8xxx_ColorSeries_for_HR.zip"
            infFile     = "BRPRC16A.INF"
        }
        3 = @{
            ip          = "192.168.2.6"
            portName    = "IP_192.168.2.6"
            printerName = "MAQ Software - Redmond IT Printer"
            driverName  = "Brother MFC-L6700DW series"
            driverURL   = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/03_MFC-L6xxx_Monochrome_for_Redmond_IT.zip"
            infFile     = "BRPRM15A.INF"
        }
        4 = @{
            ip          = "192.168.2.5"
            portName    = "IP_192.168.2.5"
            printerName = "MAQ Software - Plano Printer"
            driverName  = "Brother MFC-L3780CDW series"
            driverURL   = "https://redmonditstorage.blob.core.windows.net/redmonditpublic/Printers/Brother/01_MFC-L3xxx_ColorSeries_for_Plano.zip"
            infFile     = "BRPRC20A.INF"
        }
    }

    ########################################################################################################
    # Functions section
    ########################################################################################################

    # Function to install a specific printer by index
    function Install-PrinterByIndex {
        param ($printerIndex)
    
        $printerInfo = $script:printersTable[$printerIndex]
        Write-Log "Installing printer: $($printerInfo['printerName']) at IP: $($printerInfo['ip'])"
    
        # Remove existing printer instance
        Write-Log "Removing existing printer if it exists..."
        RemovePrinterIfExists $printerInfo['printerName']
    
        # Remove old port
        Write-Log "Removing old port if it exists..."
        DeleteOldPort $printerInfo['portName']
    
        # Setup new port
        Write-Log "Setting up new printer port: $($printerInfo['portName'])"
        SetupNewPort $printerInfo['ip'] $printerInfo['portName']
    
        # Setup Printer Driver
        Write-Log "Setting up printer driver: $($printerInfo['driverName'])"
        SetupPrinterDriver $printerInfo['driverName'] $printerInfo['driverURL'] $printerInfo['infFile']
    
        # Check if an alternative driver name was set during installation
        $finalDriverName = if ($script:actualDriverName) { $script:actualDriverName } else { $printerInfo['driverName'] }
        Write-Log "Using driver name: $finalDriverName"
    
        # Setup Printer
        Write-Log "Setting up printer with driver and port..."
        SetupPrinter $printerInfo['printerName'] $finalDriverName $printerInfo['portName']
    
        # Auto-delete downloaded files
        if ($script:downloadedDriverPath -and (Test-Path $script:downloadedDriverPath)) {
            try {
                Write-Log "Cleaning up downloaded driver files: $script:downloadedDriverPath"
                Remove-Item -Path $script:downloadedDriverPath -Recurse -Force -ErrorAction Stop
                Write-Log "Successfully cleaned up driver files"
            }
            catch {
                Write-Log "WARNING: Failed to cleanup driver files: $($_.Exception.Message)"
            }
            $script:downloadedDriverPath = $null
        }
        
        Write-Log "Completed installation of printer: $($printerInfo['printerName'])"
    }

    # Function to expand compressed driver files
    function Expand-DriverFiles {
        param ($sourcePath)
    
        $compressedFiles = Get-ChildItem $sourcePath -Recurse -File | Where-Object { $_.Extension -match '^\.\w+_$' }
        Write-Log "Found $($compressedFiles.Count) compressed files to expand"
    
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
                        }
                        catch {}
                        continue
                    }
                
                    $expandResult = & expand.exe $compFile.FullName $targetFile 2>&1
                    if ($LASTEXITCODE -eq 0 -and (Test-Path $targetFile)) {
                        $expandedCount++
                        # Remove the compressed source file after successful expansion
                        try {
                            Remove-Item $compFile.FullName -Force -ErrorAction SilentlyContinue
                        }
                        catch {}
                    }
                }
                catch {
                    # Continue on error
                }
            }
            Write-Log "Successfully expanded $expandedCount compressed driver files"
        }
    }

    # Function to delete old port
    function DeleteOldPort {
        param ($portName)
        $portExists = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
        if ($portExists) {
            try {
                Write-Log "Removing existing printer port: $portName"
                Remove-PrinterPort -Name $portName -ErrorAction Stop
                Write-Log "Successfully removed printer port: $portName"
            }
            catch {
                Write-Log "WARNING: Failed to remove printer port $portName - $($_.Exception.Message)"
            }
        } else {
            Write-Log "Port $portName does not exist, skipping removal"
        }
    }

    # Function to setup new port
    function SetupNewPort {
        param ($ip, $portName)
    
        # Check if port already exists
        $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    
        if ($existingPort) {
            Write-Log "Port $portName already exists, verifying configuration..."
            # Optionally verify the IP matches
            try {
                $cimPort = Get-CimInstance -ClassName Win32_TCPIPPrinterPort | Where-Object { $_.Name -eq $portName }
                if ($cimPort) {
                    Write-Log "Existing port verified - Name: $($cimPort.Name), IP: $($cimPort.HostAddress)"
                }
            }
            catch {
                Write-Log "Could not verify existing port details, proceeding anyway"
            }
            return
        }
    
        # Port doesn't exist, create it
        Write-Log "Creating new printer port: $portName for IP: $ip"
        # Try standard method first
        try {
            Add-PrinterPort -Name $portName -PrinterHostAddress $ip -PortNumber 9100 -ErrorAction Stop
            Write-Log "Successfully created printer port using Add-PrinterPort"
        }
        catch {
            Write-Log "Standard port creation failed, trying CIM method: $($_.Exception.Message)"
            # Fallback to CIM method
            try {
                $portProperties = @{
                    Name        = $portName
                    Protocol    = [uint32]1
                    HostAddress = $ip
                    PortNumber  = [uint32]9100
                    SNMPEnabled = $false
                }
                New-CimInstance -ClassName Win32_TCPIPPrinterPort -Property $portProperties -ErrorAction Stop | Out-Null
                Write-Log "Successfully created printer port using CIM method"
            }
            catch {
                Write-Log "ERROR: Both port creation methods failed - $($_.Exception.Message)"
                throw
            }
        }
    }

    # Function to setup printer driver
    function SetupPrinterDriver {
        param ($driverName, $driverUrl, $infFile)

        Write-Log "Setting up printer driver: $driverName from URL: $driverUrl"
        # Check if driver already exists
        $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
        if ($existingDriver) {
            Write-Log "Driver $driverName already exists, skipping installation"
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
            Write-Log "Downloading driver from: $driverUrl to: $zipFile"
            Invoke-WebRequest -Uri $driverUrl -OutFile $zipFile -ErrorAction Stop
            Write-Log "Download completed, file size: $((Get-Item $zipFile).Length) bytes"
        
            # Extract driver
            Write-Log "Extracting driver archive to: $extractPath"
            Expand-Archive -Path $zipFile -DestinationPath $extractPath -Force
            Write-Log "Extraction completed"
        
            # Find INF file - search recursively in case it's in a subfolder
            $infPath = Join-Path $extractPath $infFile
            if (-not (Test-Path $infPath)) {
                # Try searching recursively for the specific INF file
                $alternativeInf = Get-ChildItem -Path $extractPath -Filter $infFile -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($alternativeInf) {
                    $infPath = $alternativeInf.FullName
                }
                else {
                    # Try any INF file as last resort
                    $alternativeInf = Get-ChildItem "$extractPath\*.INF" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($alternativeInf) {
                        $infPath = $alternativeInf.FullName
                    }
                    else {
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
                }
                catch {
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
                    }
                    catch {
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
                }
                else {
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
            Write-Log "Creating printer: Name='$printerName', Driver='$driverName', Port='$portName'"
            Add-Printer -Name $printerName -DriverName $driverName -PortName $portName -ErrorAction Stop
            Write-Log "Successfully created printer: $printerName"
            
            # Verify printer was created
            $createdPrinter = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
            if ($createdPrinter) {
                Write-Log "Printer verification successful - Status: $($createdPrinter.PrinterStatus)"
            } else {
                Write-Log "WARNING: Printer created but not found in verification check"
            }
        }
        catch {
            Write-Log "ERROR: Failed to create printer $printerName - $($_.Exception.Message)"
            throw
        }
    }

    # Function to remove printer if it exists
    function RemovePrinterIfExists {
        param ($printerName)
        $existingPrinter = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
        if ($existingPrinter) {
            Write-Log "Removing existing printer: $printerName"
            try {
                Remove-Printer -Name $printerName -ErrorAction Stop
                Write-Log "Successfully removed printer: $printerName"
            
                # Verify removal with retry
                $retries = 0
                $maxRetries = 5
                while ($retries -lt $maxRetries) {
                    Start-Sleep -Seconds 1
                    if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                        Write-Log "Printer removal verified after $($retries + 1) attempts"
                        break
                    }
                    $retries++
                }
                
                if ($retries -ge $maxRetries) {
                    Write-Log "WARNING: Printer may not have been fully removed after $maxRetries attempts"
                }
            }
            catch {
                Write-Log "WARNING: Failed to remove existing printer $printerName - $($_.Exception.Message)"
            }
        } else {
            Write-Log "No existing printer named '$printerName' found to remove"
        }
    }

    # --- Install Printer Drivers ---
    Write-Log "Starting printer driver installation..."
    try {
        # Define printers to install (array indices: 0=Redmond Printer, 1=Redmond Color Printer, 4=Plano Printer)
        $printersToInstall = @(0, 1, 4)  # Redmond Printer, Redmond Color Printer, Plano Printer
        Write-Log "Printers to install: indices $($printersToInstall -join ', ')"
    
        foreach ($printerIndex in $printersToInstall) {
            try {
                Write-Log "Installing printer index $printerIndex..."
                Install-PrinterByIndex $printerIndex
                Write-Log "Successfully installed printer index $printerIndex"
            }
            catch {
                Write-Log "ERROR: Failed to install printer index $printerIndex - $($_.Exception.Message)"
                # Continue with next printer instead of stopping completely
                continue
            }
        }
        Write-Log "Printer driver installation section completed"
    }
    catch {
        Write-Log "ERROR: General error in printer installation section - $($_.Exception.Message)"
        # Keep downloaded files for troubleshooting
        if ($script:downloadedDriverPath -and (Test-Path $script:downloadedDriverPath)) {
            Write-Log "Printer driver files kept at: $script:downloadedDriverPath"
        }
    }
    # Mark done + delete task
    New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
    schtasks /Delete /TN "OSDCloud-DeferredSystem" /F | Out-Null

    Write-Log "=== Deferred-System completed ==="
}
finally {
    Stop-Transcript | Out-Null
}