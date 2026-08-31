#   MIT License
#
#   Copyright (c) 2026 
#
#   Łukasz Stachów
#   https://github.com/LST1618
#   https://www.linkedin.com/in/%C5%82ukasz-stach%C3%B3w-b1b886270/
#
#   Permission is hereby granted, free of charge, to any person obtaining a copy
#   of this software and associated documentation files (the "Software"), to deal
#   in the Software without restriction, including without limitation the rights
#   to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#   copies of the Software, and to permit persons to whom the Software is
#   furnished to do so, subject to the following conditions:
#
#   The above copyright notice and this permission notice shall be included in all
#   copies or substantial portions of the Software.
#
#   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#   AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#   OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#   SOFTWARE.


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