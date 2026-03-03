
$Params = @{
    OSVersion  = "Windows 11"
    OSBuild    = "25H2"
    OSEdition  = "Pro"
    OSLanguage = "en-us"
    OSLicense  = "Retail"
    ZTI        = $true
    Firmware   = $true
}

Start-OSDCloud @Params
