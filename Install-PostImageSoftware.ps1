<#
.Synopsis
    Installs post-image software packages silently and outputs installation status.
.DESCRIPTION
    This script installs various applications (e.g., Arctic Wolf, Cisco Amp, Desktop Notifier) silently,
    handles variations in installer switches and exit codes, and copies necessary configuration files.
.EXAMPLE
    .\Install-PostImageSoftware.ps1
.EXAMPLE
    Start-Process powershell -ArgumentList '-File Install-PostImageSoftware.ps1'
#>

#Script Name: Install-PostImageSoftware.ps1
#Created by Will Hughes
#Date: 2025-03-21
#Patch Notes: Initial version

# Hashtable to define silent switches for each installer
$InstallerSettings = @{
    "ArcticWolf" = @{
        Path = "\\shastahealth.org\shared\its\Store\Software\Arctic Wolf\arcticwolfagent-windows-2023-01_05-shasta0\arcticwolfagent-2023-01_05.msi"
        SilentSwitch = "/qn"   # <-- This will be updated using values from customer.json
        AcceptableExitCodes = @(0, 3010)
    }
    "CiscoAmp"   = @{
        Path = "\\shastahealth.org\shared\its\Store\Software\Cisco_Amp\Client\amp_Protect.exe"
        SilentSwitch = "/S /desktopicon 0 /startmenu 0"
        AcceptableExitCodes = @(0)
    }
    "DesktopNotifier" = @{
        Path = "\\shastahealth.org\shared\its\Store\Software\desktopnotifier\DesktopNotifier.msi"
        SilentSwitch = "/qn"
        AcceptableExitCodes = @(0, 3010)
    }
}

# For Arctic Wolf, update the silent switch using the values from customer.json.
$awCustomerJsonPath = "\\shastahealth.org\shared\its\Store\Software\Arctic Wolf\arcticwolfagent-windows-2023-01_05-shasta0\customer.json"
if (Test-Path $awCustomerJsonPath) {
    try {
        $awCustomerData = Get-Content -Path $awCustomerJsonPath | ConvertFrom-Json
        $customerUuid = $awCustomerData.customerUuid
        $registerDns  = $awCustomerData.registerDns

        # Update the silent switch string with the custom parameters and log file.
        $InstallerSettings["ArcticWolf"].SilentSwitch = "/qn CUSTOMER_UUID=$customerUuid REGISTER_DNS=$registerDns /l*v C:\InstallLogs\scout_install.log"
        Write-Host "Arctic Wolf installer parameters updated from customer.json"
    }
    catch {
        Write-Host "Error reading or parsing customer.json: $_"
    }
}
else {
    Write-Host "customer.json not found at $awCustomerJsonPath"
}

# Generic function to install a package, allowing flexibility for MSI or EXE based on file extension.
function Invoke-Install {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [Parameter(Mandatory=$true)]
        [string]$SilentSwitch,
        [Parameter(Mandatory=$true)]
        [int[]]$AcceptableExitCodes
    )
    
    Write-Host "Installing $Name from: $Path"
    # Determine the installer type based on file extension.
    $extension = [System.IO.Path]::GetExtension($Path)
    if ($extension -eq ".msi") {
        $arguments = "/i `"$Path`" $SilentSwitch"
        $process = Start-Process msiexec.exe -ArgumentList $arguments -Wait -PassThru
    }
    else {
        $process = Start-Process $Path -ArgumentList $SilentSwitch -Wait -PassThru
    }
    
    if ($AcceptableExitCodes -contains $process.ExitCode) {
        Write-Host "$Name installed successfully with exit code $($process.ExitCode)"
    }
    else {
        Write-Host "Installation of $Name failed with exit code $($process.ExitCode)"
    }
}

# Install each package using the settings defined above
foreach ($installer in $InstallerSettings.GetEnumerator()) {
    Invoke-Install -Name $installer.Key -Path $installer.Value.Path `
                   -SilentSwitch $installer.Value.SilentSwitch -AcceptableExitCodes $installer.Value.AcceptableExitCodes
}

# Copy the Spel.properties file for Desktop Notifier.
$sourceFile = "\\shastahealth.org\shared\its\Store\Software\desktopnotifier\Spel.properties"
$destinationFolder = "C:\ProgramData\Singlewire\DesktopNotifier"
$destinationFile = Join-Path $destinationFolder "Spel.properties"

Write-Host "Copying Spel.properties to $destinationFolder"
try {
    if (-not (Test-Path $destinationFolder)) {
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
    }
    Copy-Item -Path $sourceFile -Destination $destinationFile -Force
    Write-Host "Successfully copied Spel.properties"
}
catch {
    Write-Host "Failed to copy Spel.properties: $_"
}