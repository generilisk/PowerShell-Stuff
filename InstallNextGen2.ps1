# Step 1: Copy EDR Files
$sourcePathEDR = "\\ngroot\NextGenRoot\Prod\EDR"
$destinationPathEDR = "C:\Nextgen"
robocopy $sourcePathEDR $destinationPathEDR /E

# Step 2: Copy Custom Provider View
$sourcePathEHR = "\\ngroot\nextgenroot\Prod\EHR"
$destinationPathEHR = "C:\NextGen"
robocopy $sourcePathEHR $destinationPathEHR /E

# Step 3: Install Custom Font
$scriptPathFont = "\\shastahealth.org\shared\ITS\Store\Software\NextGen\scripts\NGFont\install_font.ps1"
Start-Process -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "-executionpolicy bypass -File `"$scriptPathFont`"" -Verb RunAs

# Step 4: Copy Nextgen Shortcut
$sourcePathShortcut = "\\shastahealth.org\shared\ITS\Store\Software\Shortcuts\NextGen 5.lnk"
$destinationPathShortcut = "C:\Users\Public\Desktop"
Copy-Item -Path $sourcePathShortcut -Destination $destinationPathShortcut