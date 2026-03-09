
$Params = @{
    OSVersion  = "Windows 11"
    OSBuild    = "25H2"
    OSEdition  = "Pro"
    OSLanguage = "en-us"
    OSLicense  = "Retail"
    ZTI        = $true
    Firmware   = $true
}

$Global:MyOSDCloud = @{
    WindowsUpdate          = $true
    WindowsUpdateDrivers  = $true
}


Start-OSDCloud @Params
