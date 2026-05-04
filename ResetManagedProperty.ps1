<#
    Summary: Script to reset any managed property by name, including built-in properties like Title.
             Looks up the PID from the tenant's search configuration, checks if it has been modified,
             and resets the mapping if it has.

    Usage: ResetManagedProperty.ps1 -siteUrl https://tenant-admin.sharepoint.com -managedProperty Title

    Use -printConfig to output the reset XML without applying it.
    Connect to https://tenant-admin.sharepoint.com for tenant wide reset.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$siteUrl,
    [Parameter(Mandatory = $true)]
    [string]$managedProperty,
    [string]$clientId,
    [bool]$interactiveLogin = $false,
    [bool]$printConfig = $false
)

function Load-Module ($m) {
    if (Get-Module | Where-Object { $_.Name -eq $m }) {
        write-host "Module $m is already imported."
    }
    else {
        if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $m }) {
            Import-Module $m -Verbose
        }
        else {
            if (Find-Module -Name $m | Where-Object { $_.Name -eq $m }) {
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m -Verbose
            }
            else {
                write-host "Module $m not imported, not available and not in online gallery, exiting."
                EXIT 1
            }
        }
    }
}

Load-Module "PnP.PowerShell"

if ($clientId.Length -eq 0 -and $null -eq $env:ENTRAID_CLIENT_ID -and $null -eq $env:ENTRAID_APP_ID) {
    Write-Warning "A -clientId parameter or ENTRAID_CLIENT_ID environment variable is required."
    Write-Warning "Register an app using Register-PnPEntraIDApp. See https://pnp.github.io/powershell/articles/registerapplication.html"
    return
}

$connectParams = @{ Url = $siteUrl }
if ($clientId.Length -ne 0) { $connectParams.ClientId = $clientId }
if ($interactiveLogin) { $connectParams.Interactive = $true }
Connect-PnPOnline @connectParams

# Export the current search configuration
Write-Host "Fetching search configuration from $siteUrl ..."
if ($siteUrl -like "*-admin*") {
    $searchConfig = Get-PnPSearchConfiguration -Scope Subscription
}
else {
    $searchConfig = Get-PnPSearchConfiguration -Scope Site
}

$xml = [xml]$searchConfig

$nsManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsManager.AddNamespace("d3p1", "http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration")
$nsManager.AddNamespace("d4p1", "http://schemas.microsoft.com/2003/10/Serialization/Arrays")

# Find the managed property in the Overrides section to get its PID
$overrides = $xml.SelectNodes("//d4p1:KeyValueOfstringOverrideInfoy6h3NzC8", $nsManager)

$mpPid = $null
$isModified = $false

foreach ($override in $overrides) {
    $pid = $override.SelectSingleNode("d4p1:Value/d3p1:ManagedPid", $nsManager)
    $name = $override.SelectSingleNode("d4p1:Value/d3p1:Name", $nsManager)
    $mappingsOverridden = $override.SelectSingleNode("d4p1:Value/d3p1:MappingsOverridden", $nsManager)

    if ($null -ne $name -and $name.InnerText -eq $managedProperty) {
        $mpPid = $pid.InnerText
        if ($null -ne $mappingsOverridden -and $mappingsOverridden.InnerText -eq "true") {
            $isModified = $true
        }
        break
    }
}

# If not found in overrides, look in the ManagedProperties section
if ($null -eq $mpPid) {
    $managedProps = $xml.SelectNodes("//d4p1:KeyValueOfstringManagedPropertyInfoy6h3NzC8", $nsManager)
    foreach ($mp in $managedProps) {
        $key = $mp.SelectSingleNode("d4p1:Key", $nsManager)
        if ($null -ne $key -and $key.InnerText -eq $managedProperty) {
            $pid = $mp.SelectSingleNode("d4p1:Value/d3p1:Pid", $nsManager)
            if ($null -ne $pid) {
                $mpPid = $pid.InnerText
            }
            break
        }
    }
}

if ($null -eq $mpPid) {
    Write-Warning "Managed property '$managedProperty' not found in the search configuration."
    Write-Warning "Make sure you are connected to the correct scope (tenant admin for Subscription, site for Site)."
    return
}

Write-Host "Found managed property '$managedProperty' with PID: $mpPid"

if (-not $isModified) {
    Write-Host "Managed property '$managedProperty' has not been modified (MappingsOverridden=false). No reset needed."
    return
}

Write-Host "Managed property '$managedProperty' has been modified. Resetting..."

# Load the reset template and substitute the PID
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
if ($scriptDir.Length -eq 0) {
    $scriptDir = "."
}

$config = Get-Content -Path "$scriptDir\SearchMappingReset.xml" -Raw
$config = $config -replace "##PID##", $mpPid

if ($printConfig) {
    $config
    return
}

if ($siteUrl -like "*-admin*") {
    Set-PnPSearchConfiguration -Scope Subscription -Configuration $config
}
else {
    Set-PnPSearchConfiguration -Scope Site -Configuration $config
}

Write-Host "Reset complete for '$managedProperty'. Remember to re-index items after a schema change."
