<#
    Author: Mikael Svenson - techmikael.com - @mikaelsvenson
    Date: November 2020

    Summary: Script to map or clear crawled properties to reusable managed properties

    Usage: MapCrawledPropertyToManagedProperty.ps1 -siteUrl https://tenant.sharepoint.com/sites/site -managedProperty RefinableString00 -crawledProperty ows_<columnName>

    Clear/reset a mapping with: MapCrawledPropertyToManagedProperty.ps1 -siteUrl https://tenant.sharepoint.com/sites/site -managedProperty RefinableString00

    Use -appendToExistingMapping:$false if you want to overwrite an existing mapping instead of adding to it.
    Connect to https://tenant-admin.sharepoint.com for tenant wide mapping
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$siteUrl,
    [Parameter(Mandatory = $true)]
    [string]$managedProperty,    
    [string]$crawledProperty,
    [string]$clientId,
    [bool]$appendToExistingMapping = $true,
    [bool]$interactiveLogin = $false,
    [bool]$printConfig = $false
)


function Import-RequiredModule ($m) {
    # If module is imported say that and do nothing
    if (Get-Module | Where-Object { $_.Name -eq $m }) {
        write-host "Module $m is already imported."
    }
    else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $m }) {
            Import-Module $m -Verbose
        }
        else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object { $_.Name -eq $m }) {
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m -Verbose
            }
            else {

                # If module is not imported, not available and not in online gallery then abort
                write-host "Module $m not imported, not available and not in online gallery, exiting."
                EXIT 1
            }
        }
    }
}

function Get-LegacyManagedPropertyPid ([string]$managedPropertyName, [string]$mpNumber) {
    $basePid = 0;
    $legacyMpNumber = $mpNumber

    if ($managedPropertyName -match "^RefinableString1" -and $legacyMpNumber.length -eq 3 ) {
        $basePid = 1000000900
        $legacyMpNumber = $legacyMpNumber - 100;
    }
    elseif ($managedPropertyName -match "^RefinableDouble") {
        $basePid = 1000000800
    }
    elseif ($managedPropertyName -match "^RefinableDecimal") {
        $basePid = 1000000700
    }
    elseif ($managedPropertyName -match "^RefinableDateInvariant") {
        $basePid = 1000000660
    }
    elseif ($managedPropertyName -match "^RefinableDateSingle") {
        $basePid = 1000000650
    }
    elseif ($managedPropertyName -match "^RefinableDate") {
        $basePid = 1000000600
    }
    elseif ($managedPropertyName -match "^RefinableInt") {
        $basePid = 1000000500
    }
    elseif ($managedPropertyName -match "^Double") {
        $basePid = 1000000400
    }
    elseif ($managedPropertyName -match "^Decimal") {
        $basePid = 1000000300
    }
    elseif ($managedPropertyName -match "^Date") {
        $basePid = 1000000200
    }
    elseif ($managedPropertyName -match "^Int") {
        $basePid = 1000000100
    }
    elseif ($managedPropertyName -match "^RefinableString") {
        $basePid = 1000000000
    }

    return ($basePid + [int]$legacyMpNumber)
}

function Get-SearchSchemaContext ([string]$managedPropertyName, [string]$scope) {
    $result = [PSCustomObject]@{
        ManagedPropertyPid = $null
        SchemaId = $null
    }

    try {
        if ($scope -eq "Subscription") {
            $searchConfig = Get-PnPSearchConfiguration -Scope Subscription
        }
        else {
            $searchConfig = Get-PnPSearchConfiguration -Scope Site
        }

        $xml = [xml]$searchConfig

        $schemaIdNode = $xml.SelectSingleNode("//*[local-name()='SchemaId']")
        if ($null -ne $schemaIdNode -and -not [string]::IsNullOrWhiteSpace($schemaIdNode.InnerText)) {
            $result.SchemaId = $schemaIdNode.InnerText
        }

        $nsManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $nsManager.AddNamespace("d3p1", "http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration")
        $nsManager.AddNamespace("d4p1", "http://schemas.microsoft.com/2003/10/Serialization/Arrays")

        $managedProps = $xml.SelectNodes("//d4p1:KeyValueOfstringManagedPropertyInfoy6h3NzC8", $nsManager)
        foreach ($mp in $managedProps) {
            $key = $mp.SelectSingleNode("d4p1:Key", $nsManager)
            if ($null -ne $key -and $key.InnerText -eq $managedPropertyName) {
                $pid = $mp.SelectSingleNode("d4p1:Value/d3p1:Pid", $nsManager)
                if ($null -ne $pid) {
                    $result.ManagedPropertyPid = [int]$pid.InnerText
                    return $result
                }
            }
        }

        $overrides = $xml.SelectNodes("//d4p1:KeyValueOfstringOverrideInfoy6h3NzC8", $nsManager)
        foreach ($override in $overrides) {
            $name = $override.SelectSingleNode("d4p1:Value/d3p1:Name", $nsManager)
            if ($null -ne $name -and $name.InnerText -eq $managedPropertyName) {
                $managedPid = $override.SelectSingleNode("d4p1:Value/d3p1:ManagedPid", $nsManager)
                if ($null -ne $managedPid) {
                    $result.ManagedPropertyPid = [int]$managedPid.InnerText
                    return $result
                }
            }
        }
    }
    catch {
        Write-Warning "Could not resolve PID from search configuration: $($_.Exception.Message)"
    }

    return $result
}

Import-RequiredModule "PnP.PowerShell"

$hasClientId = -not [string]::IsNullOrWhiteSpace($clientId)
$hasEntraClientId = -not [string]::IsNullOrWhiteSpace($env:ENTRAID_CLIENT_ID)
$hasEntraAppId = -not [string]::IsNullOrWhiteSpace($env:ENTRAID_APP_ID)

if (-not $hasClientId -and -not $hasEntraClientId -and -not $hasEntraAppId) {
    Write-Warning "A -clientId parameter or ENTRAID_CLIENT_ID environment variable is required."
    Write-Warning "Register an app using Register-PnPEntraIDApp. See https://pnp.github.io/powershell/articles/registerapplication.html"
    return
}

$connectParams = @{ Url = $siteUrl }
if ($hasClientId) { $connectParams.ClientId = $clientId }
if ($interactiveLogin) { $connectParams.Interactive = $true }
Connect-PnPOnline @connectParams

$searchScope = "Site"
if ($siteUrl -like "*-admin*") {
    $searchScope = "Subscription"
}

$hasCrawledProperty = -not [string]::IsNullOrWhiteSpace($crawledProperty)

if ($hasCrawledProperty -and -not ($crawledProperty -match '^ows_')) {
    Write-Warning "Crawled property has to start with ows_";
    return
}

if ($hasCrawledProperty -and ($crawledProperty -match '^(ows_taxId|ows_r_|ows_q_)')) {
    Write-Warning "Script only support regular crawled properties ows_<field name>";
    return
}

$mp = $managedProperty  | Select-String -Pattern '^(?<name>(Refinable|Double|Decimal|Date|Int)[a-z]*)(?<num>\d+)$'

if ($mp.Matches.Length -eq 0) {
    Write-Warning "Script only support reusable managed properties";
    return
}

$mpNumber = $mp.Matches[0].Groups['num'].Value

$searchSchemaContext = Get-SearchSchemaContext -managedPropertyName $managedProperty -scope $searchScope

$mpPid = $searchSchemaContext.ManagedPropertyPid
if ($null -eq $mpPid) {
    Write-Warning "Could not find '$managedProperty' in live search configuration. Falling back to legacy PID formula."
    $mpPid = Get-LegacyManagedPropertyPid -managedPropertyName $managedProperty -mpNumber $mpNumber
    Write-Host "Using fallback PID $mpPid for managed property '$managedProperty'"
}
else {
    Write-Host "Using resolved PID $mpPid for managed property '$managedProperty'"
}

$schemaId = $searchSchemaContext.SchemaId
if ([string]::IsNullOrWhiteSpace($schemaId)) {
    $schemaId = "143692"
    Write-Warning "Could not resolve SchemaId from live search configuration. Falling back to $schemaId"
}
else {
    Write-Host "Using resolved SchemaId $schemaId"
}

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
if ($scriptDir.Length -eq 0) {
    $scriptDir = "."
}
if (-not $hasCrawledProperty) {
    Write-Host "Resetting crawled property mapping for $managedProperty"
    $config = Get-Content -Path "$scriptDir\SearchMappingReset.xml" -Raw
}
else {
    if ($appendToExistingMapping) {
        Write-Host "Appending crawled property $crawledProperty to managed property $managedProperty"
    }
    else {
        Write-Host "Replacing crawled property $crawledProperty to managed property $managedProperty"
    }    
    $config = Get-Content -Path "$scriptDir\SearchMappingTemplate.xml" -Raw
}

$config = $config -replace "##PID##", $mpPid
$config = $config -replace "##CPNAME##", $crawledProperty
$config = $config -replace "##APPEND##", $appendToExistingMapping.ToString().ToLower()
$config = $config -replace "##SCHEMAID##", $schemaId

if ($printConfig) {
    $config
    return
}

if ($searchScope -eq "Subscription") {
    Set-PnPSearchConfiguration -Scope Subscription -Configuration $config
}
else {
    Set-PnPSearchConfiguration -Scope Site -Configuration $config
}

Write-Host "Remember to re-index items after you have done a schema change"



