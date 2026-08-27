# Export-ADSchemaObjectCategory.ps1
# Outputs all schema classes, their defaultObjectCategory, and ready-to-use LDAP filters into excel and .json format

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputFolder = "C:\Temp\ADSchemaExport"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Required Modules 
#Import-Module ActiveDirectory
#Import-Module ImportExcel

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$rootDse = Get-ADRootDSE
$schemaNC = $rootDse.schemaNamingContext

$schemaClasses = Get-ADObject `
    -SearchBase $schemaNC `
    -LDAPFilter '(objectClass=classSchema)' `
    -Properties lDAPDisplayName, defaultObjectCategory, subClassOf, governsID, adminDisplayName, description |
    Sort-Object lDAPDisplayName |
    Select-Object `
        lDAPDisplayName,
        defaultObjectCategory,
        subClassOf,
        governsID,
        adminDisplayName,
        description,
        @{Name='LDAPFilterExample';Expression={
            if ($_.lDAPDisplayName) {
                "(objectCategory=$($_.lDAPDisplayName))"
            }
        }},
        @{Name='RawDefaultObjectCategoryFilter';Expression={
            if ($_.defaultObjectCategory) {
                "(objectCategory=$($_.defaultObjectCategory))"
            }
        }}


$xlsxPath = Join-Path $OutputFolder 'SchemaObjectCategoryMap.xlsx'
$jsonPath = Join-Path $OutputFolder 'SchemaObjectCategoryMap.json'


$schemaClasses | ConvertTo-Json -Depth 5 | Set-Content -Path $jsonPath -Encoding UTF8

if (Get-Module -ListAvailable -Name ImportExcel) {
    
    $schemaClasses | Export-Excel -Path $xlsxPath -WorksheetName 'SchemaMap' -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium2
}
else {
    Write-Warning "ImportExcel module not found, so XLSX was not created."
}

$schemaClasses