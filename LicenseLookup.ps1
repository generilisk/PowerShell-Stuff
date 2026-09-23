<#
.SYNOPSIS
Shared Microsoft Graph license-lookup helpers, used by both Get-UserLicenses.ps1
and Update-OffboardList.ps1.

.DESCRIPTION
This file is not meant to be run directly. Dot-source it from another script:

    . "$PSScriptRoot\LicenseLookup.ps1"

It provides:
- $skuFriendlyNames       : hashtable mapping raw SKU part numbers to friendly product names
- Get-FriendlyName        : looks up a friendly name for a raw SKU part number
- Connect-LicenseGraph    : wraps Connect-MgGraph with consistent scope/error handling
- Get-TenantSkuCache      : wraps Get-MgSubscribedSku -All (call once, reuse across users)
- Resolve-UserLicense     : given a username/UPN, DefaultDomain, and the cached SKU list,
                            returns a friendly license summary string for that user.

Keep all SKU-name mappings and lookup logic here. If Microsoft adds new SKUs to your
tenant, or the lookup strategy needs to change, update it once in this file and both
scripts pick up the change automatically.

Connect-LicenseGraph also installs Microsoft.Graph.Users and
Microsoft.Graph.Identity.DirectoryManagement (CurrentUser scope) on first
use if either is missing, so a fresh machine doesn't need manual module
setup before running either script.
#>

# --- Friendly name lookup table (extend as needed for your tenant) ---
# Full official list: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
$skuFriendlyNames = @{
    "ENTERPRISEPACK"            = "Office 365 E3"
    "ENTERPRISEPREMIUM"         = "Office 365 E5"
    "SPE_E3"                    = "Microsoft 365 E3"
    "SPE_E5"                    = "Microsoft 365 E5"
    "SPB"                       = "Microsoft 365 Business Premium"
    "O365_BUSINESS_ESSENTIALS"  = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"     = "Microsoft 365 Business Standard"
    "EXCHANGESTANDARD"          = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE"        = "Exchange Online Plan 2"
    "AAD_PREMIUM"               = "Entra ID P1"
    "AAD_PREMIUM_P2"            = "Entra ID P2"
    "FLOW_FREE"                 = "Power Automate Free"
    "POWER_BI_STANDARD"         = "Power BI (Free)"
    "TEAMS_EXPLORATORY"         = "Teams Exploratory"
}

function Get-FriendlyName {
    param([string]$SkuPartNumber)
    if ($skuFriendlyNames.ContainsKey($SkuPartNumber)) {
        return $skuFriendlyNames[$SkuPartNumber]
    }
    return $SkuPartNumber  # fall back to raw name if not mapped
}

function Connect-LicenseGraph {
    <#
    .SYNOPSIS
    Ensures required Graph modules are installed, then connects to Microsoft Graph
    with the scope needed for license lookups.
    Returns $true on success, $false on failure (and writes an error).
    #>
    foreach ($moduleName in @('Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement')) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            try {
                Write-Host "Installing required module: $moduleName..." -ForegroundColor Cyan
                Install-Module -Name $moduleName -Scope CurrentUser -Force -ErrorAction Stop
            }
            catch {
                Write-Error "Failed to install required module '$moduleName': $_"
                return $false
            }
        }
    }

    try {
        Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
        return $true
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        return $false
    }
}

function Get-TenantSkuCache {
    <#
    .SYNOPSIS
    Fetches the tenant's subscribed SKUs once. Call this a single time per script run
    and pass the result into Resolve-UserLicense for every user, rather than
    re-querying per user/license.
    #>
    return Get-MgSubscribedSku -All
}

function Resolve-UserLicense {
    <#
    .SYNOPSIS
    Resolves a username/UPN to its identity and assigned licenses in one call.
    This is the single place the UPN/SAM-account matching strategy lives -
    callers should not re-implement it themselves.

    .PARAMETER Username
    A bare username (e.g. "atippens"), a full UPN (e.g. "atippens@shastahealth.org"),
    or an Entra object GUID.

    .PARAMETER DefaultDomain
    Domain to append to bare usernames when constructing a UPN, e.g. "shastahealth.org".
    Leave blank to skip straight to the onPremisesSamAccountName filter fallback.

    .PARAMETER TenantSkus
    The cached result of Get-TenantSkuCache, passed in so it isn't re-fetched per user.

    .OUTPUTS
    A PSCustomObject with:
        DisplayName        - resolved display name, or "N/A" if not found
        UserPrincipalName  - resolved UPN, or the original input if not found
        Licenses           - semicolon-joined friendly license names, "(none assigned)",
                              "N/A - username not resolved" (blank input), or "ERROR: <message>"
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Username,

        [string]$DefaultDomain,

        [Parameter(Mandatory)]
        $TenantSkus
    )

    $u = $Username.Trim()

    if ([string]::IsNullOrWhiteSpace($u)) {
        return [PSCustomObject]@{
            DisplayName       = "N/A"
            UserPrincipalName = $u
            Licenses          = "N/A - username not resolved"
        }
    }

    try {
        $mgUser = $null

        # If it already looks like a UPN (contains @) or a GUID, try it directly first
        if ($u -match "@" -or $u -match "^[0-9a-fA-F-]{36}$") {
            try {
                $mgUser = Get-MgUser -UserId $u -Property "DisplayName,UserPrincipalName,AssignedLicenses" -ErrorAction Stop
            }
            catch { $mgUser = $null }
        }

        # Try constructed UPN using DefaultDomain
        if (-not $mgUser -and $DefaultDomain -and $u -notmatch "@") {
            try {
                $constructedUpn = "$u@$DefaultDomain"
                $mgUser = Get-MgUser -UserId $constructedUpn -Property "DisplayName,UserPrincipalName,AssignedLicenses" -ErrorAction Stop
            }
            catch { $mgUser = $null }
        }

        # Fall back to filtering by on-prem SAM account name (works regardless of UPN format)
        if (-not $mgUser) {
            $filterMatch = Get-MgUser -Filter "onPremisesSamAccountName eq '$u'" -Property "DisplayName,UserPrincipalName,AssignedLicenses" -ErrorAction Stop
            if ($filterMatch) {
                $mgUser = $filterMatch | Select-Object -First 1
            }
        }

        if (-not $mgUser) {
            throw "No matching user found by UPN or onPremisesSamAccountName"
        }

        if ($mgUser.AssignedLicenses.Count -eq 0) {
            $licenseText = "(none assigned)"
        }
        else {
            $licenseNames = foreach ($lic in $mgUser.AssignedLicenses) {
                $sku = $TenantSkus | Where-Object { $_.SkuId -eq $lic.SkuId }
                if ($sku) { Get-FriendlyName $sku.SkuPartNumber } else { $lic.SkuId }
            }
            $licenseText = ($licenseNames -join "; ")
        }

        return [PSCustomObject]@{
            DisplayName       = $mgUser.DisplayName
            UserPrincipalName = $mgUser.UserPrincipalName
            Licenses          = $licenseText
        }
    }
    catch {
        return [PSCustomObject]@{
            DisplayName       = "N/A"
            UserPrincipalName = $u
            Licenses          = "ERROR: $($_.Exception.Message)"
        }
    }
}
