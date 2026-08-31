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



[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath = "C:\Temp\ADObjects.xlsx"
)


Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

### Modules requried to run this script
#Install-Module -Name ImportExcel -Force
#Import-Module ActiveDirectory
#Import-Module ImportExcel

$users = Get-ADUser -Filter * -Properties DisplayName,Department,Title,EmailAddress,Enabled,DistinguishedName |
    Select-Object DisplayName,SamAccountName,UserPrincipalName,Department,Title,EmailAddress,Enabled,DistinguishedName

$groups = Get-ADGroup -Filter * -Properties GroupScope,GroupCategory,Description,DistinguishedName |
    Select-Object Name,SamAccountName,GroupScope,GroupCategory,Description,DistinguishedName

$computers = Get-ADComputer -Filter * -Properties DNSHostName,OperatingSystem,Enabled,DistinguishedName |
    Select-Object Name,SamAccountName,DNSHostName,OperatingSystem,Enabled,DistinguishedName

$ous = Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName,Description |
    Select-Object Name,DistinguishedName,Description


### Remaining Objects from AD could be done in a similar way.
### Below also an example on how to use LDAPfilter to get specific objects from AD.
### Related Articles:
### https://learn.microsoft.com/en-us/windows/win32/adsi/objectcategory-vs--objectclass
### Full Export of All LDAPfilter searchable values in the other Export-ADSchemaObjectCategory.ps1 script.
### $printers = Get-ADObject -LDAPFilter '(objectCategory=printQueue)' -Properties *
### $person = Get-ADObject -LDAPFilter '(objectCategory=person)' -Properties *



$users     | Export-Excel -Path $WorkbookPath -WorksheetName 'Users'     -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet -TableStyle Medium2
$groups    | Export-Excel -Path $WorkbookPath -WorksheetName 'Groups'    -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet -TableStyle Medium2
$computers | Export-Excel -Path $WorkbookPath -WorksheetName 'Computers' -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet -TableStyle Medium2
$ous       | Export-Excel -Path $WorkbookPath -WorksheetName 'OUs'       -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet -TableStyle Medium2

### $person       | Export-Excel -Path $WorkbookPath -WorksheetName 'Persons'       -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet -TableStyle Medium2