

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