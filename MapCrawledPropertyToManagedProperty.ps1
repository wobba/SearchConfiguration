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
    [bool]$appendToExistingMapping = $true
)


function Load-Module ($m) {
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

Load-Module "SharePointPnPPowerShellOnline"

Connect-PnPOnline -Url $siteUrl -UseWebLogin

$validNames = @("Int", "Date", "Decimal", "Double", "RefinableInt", "RefinableDate", "RefinableDateSingle", "RefinableDateInvariant", "RefinableDecimal", "RefinableDecimal", "RefinableString");

if (($crawledProperty.Length -ne 0) -and -not ($crawledProperty -match '^ows_')) {
    Write-Warning "Crawled property has to start with ows_";
    return
}

if ($crawledProperty -match '^(ows_taxId|ows_r|ows_q)') {
    Write-Warning "Script only support regular crawled properties ows_<field name>";
    return
}

$mp = $managedProperty  | Select-String -Pattern '^(?<name>(Refinable|Double|Decimal|Date|Int)[a-z]*)(?<num>\d+)$'

if ($mp.Matches.Length -eq 0) {
    Write-Warning "Script only support reusable managed properties";
    return
}

$mpNumber = $mp.Matches[0].Groups['num'].Value

$basePid = 0;
if ($managedProperty -match "^RefinableString1") {
    $basePid = 1000000900
    $mpNumber = $mpNumber - 100;
}
elseif ($managedProperty -match "^RefinableDouble") {
    $basePid = 1000000800
}
elseif ($managedProperty -match "^RefinableDecimal") {
    $basePid = 1000000700
}
elseif ($managedProperty -match "^RefinableDateInvariant") {
    $basePid = 1000000660
}
elseif ($managedProperty -match "^RefinableDateSingle") {
    $basePid = 1000000660
}
elseif ($managedProperty -match "^RefinableDate") {
    $basePid = 1000000600
}
elseif ($managedProperty -match "^RefinableInt") {
    $basePid = 1000000500
}
elseif ($managedProperty -match "^Double") {
    $basePid = 1000000400
}
elseif ($managedProperty -match "^Decimal") {
    $basePid = 1000000300
}
elseif ($managedProperty -match "^Date") {
    $basePid = 1000000200
}
elseif ($managedProperty -match "^Int") {
    $basePid = 1000000100
}
elseif ($managedProperty -match "^RefinableString") {
    $basePid = 1000000000
}

$mpPid = $basePid + [int]$mpNumber

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
if ($scriptDir.Length -eq 0) {
    $scriptDir = "."
}
if ($crawledProperty.Length -eq 0) {
    Write-Host "Resetting crawled property mapping for $managedProperty"
    $config = Get-Content -Path "$scriptDir\SearchMappingReset.xml" -Raw
}
else {
    if($appendToExistingMapping) {
        Write-Host "Appending crawled property $crawledProperty to managed property $managedProperty"
    } else {
        Write-Host "Replacing crawled property $crawledProperty to managed property $managedProperty"
    }    
    $config = Get-Content -Path "$scriptDir\SearchMappingTemplate.xml" -Raw
}

$config = $config -replace "##PID##", $mpPid
$config = $config -replace "##CPNAME##", $crawledProperty
$config = $config -replace "##APPEND##", $appendToExistingMapping.ToString().ToLower()

if ($siteUrl -like "*-admin*") {
    Set-PnPSearchConfiguration -Scope Subscription -Configuration $config
}
else {
    Set-PnPSearchConfiguration -Scope Site -Configuration $config
}

Write-Host "Remember to re-index items after you have done a schema change"



