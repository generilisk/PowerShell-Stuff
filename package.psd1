@{
    Root = 'h:\PowerShell\VSCode\PowerShell-Stuff-1\PrinterManager.ps1'
    OutputPath = 'h:\PowerShell\VSCode\PowerShell-Stuff-1\out'
    Package = @{
        Enabled = $true
        Obfuscate = $false
        HideConsoleWindow = $true
        DotNetVersion = 'v4.6.2'
        FileVersion = '1.0.0'
        FileDescription = 'Provide a GUI for personal printer management'
        ProductName = 'SCHC Printer Manager'
        ProductVersion = '1.0.0'
        Copyright = 'Copyright © 2024 Will Hughes. All rights reserved.'
        RequireElevation = $false
        ApplicationIconPath = 'C:\Users\whughes\Downloads\Oxygen-Icons.org-Oxygen-Devices-printer-laser.ico'
        PackageType = 'Console'
    }
    Bundle = @{
        Enabled = $true
        Modules = $true
        # IgnoredModules = @()
    }
}
        