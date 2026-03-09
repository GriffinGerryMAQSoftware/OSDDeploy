Start-Transcript C:\Windows\Temp\SetupComplete-OSDCloud.log

# Ensure working folders exist
New-Item -ItemType Directory -Path C:\ProgramData\OSDCloud\Assets  -Force
New-Item -ItemType Directory -Path C:\ProgramData\OSDCloud\Scripts -Force

# Copy assets staged by OSDCloud
Copy-Item C:\OSDCloud\Config\Assets\*  C:\ProgramData\OSDCloud\Assets  -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item C:\OSDCloud\Config\Scripts\oobe.ps1 C:\ProgramData\OSDCloud\Scripts\oobe.ps1 -Force

# Run your actual configuration
& C:\ProgramData\OSDCloud\Scripts\oobe.ps1

Stop-Transcript