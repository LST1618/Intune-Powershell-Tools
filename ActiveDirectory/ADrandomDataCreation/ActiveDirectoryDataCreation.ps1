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


<#
Lab Active Directory
Creates:
- OU structure for Entra Connect and AD training
- ~240 users by default
- ~60 computer accounts by default
- Security/distribution groups
- Contacts
- Service/admin/test/shared/resource accounts
- Disabled users and stale-style objects
- Manager relationships and common synced attributes

Run as Domain Admin or delegated account with rights to create objects.
Requires: ActiveDirectory module
Recommended: run only in a lab / training domain
This version without Exchange attributes
#>

[CmdletBinding()]
param(
    [int]$UserCount = 220,
    [int]$ComputerCount = 60,
    [string]$CompanyName = "WhateverCompany",
    [string]$UPNSuffix = "",
    [string]$DefaultPassword = "P@ssw0rd",
    [switch]$WhatIfMode
)

$ErrorActionPreference = "Stop"

if ($WhatIfMode) { $WhatIfPreference = $true }

#Import-Module ActiveDirectory

# ----------------------------
# Helpers
# ----------------------------
function Write-Log {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function New-LabOU {
    param(
        [string]$Name,
        [string]$Path,
        [bool]$Protected = $false
    )
    $existing = Get-ADOrganizationalUnit -LDAPFilter "(ou=$Name)" -SearchBase $Path -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-ADOrganizationalUnit -Name $Name -Path $Path -ProtectedFromAccidentalDeletion:$Protected
        Write-Log "Created OU: OU=$Name,$Path"
    }
}

function Ensure-Group {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Scope = "Global",
        [string]$Category = "Security",
        [string]$Description = ""
    )
    $g = Get-ADGroup -LDAPFilter "(cn=$Name)" -SearchBase $Path -ErrorAction SilentlyContinue
    if (-not $g) {
        New-ADGroup -Name $Name -SamAccountName $Name -GroupScope $Scope -GroupCategory $Category -Path $Path -Description $Description
        Write-Log "Created group: $Name"
    }
}

function Get-UniqueSam {
    param(
        [string]$BaseSam,
        [string]$SearchBase
    )
    $sam = $BaseSam.ToLower()
    $i = 1
    while (Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" -SearchBase $SearchBase -ErrorAction SilentlyContinue) {
        $sam = ($BaseSam.Substring(0, [Math]::Min(16, $BaseSam.Length)) + $i).ToLower()
        $i++
    }
    return $sam.Substring(0, [Math]::Min(20, $sam.Length))
}

function Get-UPNSuffix {
    param([string]$FallbackDnsRoot)
    if ($UPNSuffix -and $UPNSuffix.Trim()) { return $UPNSuffix.Trim() }
    try {
        $forest = Get-ADForest
        if ($forest.UPNSuffixes.Count -gt 0) { return $forest.UPNSuffixes[0] }
    } catch {}
    return $FallbackDnsRoot
}

function Set-OptionalUserAttributes {
    param(
        [string]$Identity,
        [hashtable]$Replace
    )
    $clean = @{}
    foreach ($k in $Replace.Keys) {
        if ($null -ne $Replace[$k] -and "$($Replace[$k])".Trim() -ne "") {
            $clean[$k] = $Replace[$k]
        }
    }
    if ($clean.Count -gt 0) {
        Set-ADUser -Identity $Identity -Replace $clean
    }
}

# ----------------------------
# Domain context
# ----------------------------
$domain = Get-ADDomain
$rootDN = $domain.DistinguishedName
$dnsRoot = $domain.DNSRoot
$netbios = $domain.NetBIOSName
$effectiveUPNSuffix = Get-UPNSuffix -FallbackDnsRoot $dnsRoot
$securePassword = ConvertTo-SecureString $DefaultPassword -AsPlainText -Force

Write-Log "Domain DN: $rootDN"
Write-Log "DNS Root: $dnsRoot"
Write-Log "Using UPN suffix: $effectiveUPNSuffix"

# ----------------------------
# Top-level OU design
# ----------------------------
$topOU = "OU=LAB,$rootDN"
if (-not (Get-ADOrganizationalUnit -LDAPFilter "(ou=LAB)" -SearchBase $rootDN -ErrorAction SilentlyContinue)) {
    New-ADOrganizationalUnit -Name "LAB" -Path $rootDN -ProtectedFromAccidentalDeletion:$true
    Write-Log "Created top OU: LAB"
}

# Core OUs
$ouList = @(
    @{ Name = "Users"; Path = $topOU },
    @{ Name = "Computers"; Path = $topOU },
    @{ Name = "Groups"; Path = $topOU },
    @{ Name = "Contacts"; Path = $topOU },
    @{ Name = "Service Accounts"; Path = $topOU },
    @{ Name = "Admins"; Path = $topOU },
    @{ Name = "Disabled Objects"; Path = $topOU },
    @{ Name = "NonSync"; Path = $topOU },
    @{ Name = "Resources"; Path = $topOU },
    @{ Name = "EntraPilot"; Path = $topOU }
)

foreach ($ou in $ouList) {
    New-LabOU -Name $ou.Name -Path $ou.Path -Protected $false
}

# User sub-OUs
$userRoot = "OU=Users,$topOU"
$userOUs = @("HQ","Krakow","Wroclaw","Poznan","Gdansk","Sales","IT","HR","Finance","Operations","Interns","Marketing","Support","Engineering")
foreach ($ou in $userOUs) {
    New-LabOU -Name $ou -Path $userRoot
}

# Computer sub-OUs
$computerRoot = "OU=Computers,$topOU"
$computerOUs = @("Workstations","Laptops","Kiosks","Servers","Legacy","PilotDevices")
foreach ($ou in $computerOUs) {
    New-LabOU -Name $ou -Path $computerRoot
}

# Group sub-OUs
$groupRoot = "OU=Groups,$topOU"
$groupOUs = @("Security","Distribution","Licensing","Applications","RBAC")
foreach ($ou in $groupOUs) {
    New-LabOU -Name $ou -Path $groupRoot
}

# Resource sub-OUs
$resourceRoot = "OU=Resources,$topOU"
$resourceOUs = @("Shared Mailboxes","Rooms","Project Accounts")
foreach ($ou in $resourceOUs) {
    New-LabOU -Name $ou -Path $resourceRoot
}

# NonSync sub-OUs
$nonSyncRoot = "OU=NonSync,$topOU"
$nonSyncOUs = @("Users","Service Accounts","Computers")
foreach ($ou in $nonSyncOUs) {
    New-LabOU -Name $ou -Path $nonSyncRoot
}

# Disabled sub-OUs
$disabledRoot = "OU=Disabled Objects,$topOU"
$disabledOUs = @("Users","Computers")
foreach ($ou in $disabledOUs) {
    New-LabOU -Name $ou -Path $disabledRoot
}

# Entra pilot sub-OUs
$pilotRoot = "OU=EntraPilot,$topOU"
$pilotOUs = @("Users","Devices","Groups")
foreach ($ou in $pilotOUs) {
    New-LabOU -Name $ou -Path $pilotRoot
}

# ----------------------------
# Data pools
# ----------------------------
$firstNames = @(
    "Adam","Adrian","Agnieszka","Aleksandra","Andrzej","Anna","Bartosz","Beata","Cezary","Damian","Daniel","Daria",
    "Dawid","Dominik","Ewa","Filip","Grzegorz","Hubert","Iga","Izabela","Jakub","Jan","Joanna","Kacper","Karol",
    "Katarzyna","Konrad","Krzysztof","Laura","Lena","Lukasz","Magda","Maja","Marek","Marcin","Mateusz","Michal",
    "Monika","Natalia","Nikodem","Olga","Patryk","Pawel","Piotr","Rafal","Sandra","Sebastian","Sylwia","Tomasz",
    "Weronika","Wiktor","Zofia","Alicja","Antoni","Barbara","Blazej","Bogdan","Bozena","Celina","Cyprian","Dariusz","Dorota","Edyta","Emilia",
    "Eryk","Fabian","Felicja","Gabriel","Genowefa","Gustaw","Halina","Henryk","Ida","Ignacy","Irena","Ireneusz",
    "Jadwiga","Jaroslaw","Jerzy","Julia","Julian","Justyna","Kamil","Kamila","Karolina","Klara","Klaudia","Krystian",
    "Krystyna","Ksawery","Leon","Leszek","Liwia","Lidia","Malgorzata","Marta","Marian","Mariusz","Marzena","Mikolaj",
    "Milena","Miroslaw","Nadia","Nina","Norbert","Oliwia","Oskar","Robert","Roman","Roza","Ryszard","Slawomir",
    "Stanislaw","Stefan","Szymon","Teodor","Teresa","Urszula","Wanda","Waldemar","Wieslaw","Wladyslaw","Zbigniew",
    "Zdzislaw","Zenon","Zuzanna","Aleksander","Amelia","Aniela","Bogumila","Bronislaw","Czeslaw","Elzbieta","Feliks",
    "Florentyna","Franciszek","Gabriela","Gerard","Grazyna","Iwona","Jacek","Kazimiera","Kordian","Lucja","Marlena",
    "Natan","Radoslaw","Wojciech","Zdzislawa"
)

$lastNames = @(
    "Nowak","Kowalski","Wisniewski","Wojcik","Kowalczyk","Kaminski","Lewandowski","Zielinski","Szymanski","Wozniak",
    "Dabrowski","Kozlowski","Jankowski","Mazur","Krawczyk","Piotrowski","Grabowski","Nowicka","Pawlak","Michalski",
    "Krol","Wieczorek","Jablonski","Wrobel","Stepien","Duda","Zajac","Adamczyk","Pawlowski","Walczak",
    "Sikora","Baran","Rutkowski","Michalak","Szewczyk","Ostrowski","Tomaszewski","Pietrzak","Wojciechowski","Jaworski",
    "Wojtowicz","Sadowski","Marciniak","Zawadzki","Sikorski","Wysocki","Wieczorkowski","Urbaniak","Kwiatkowski","Wilczek",
    "Kubiak","Kaczmarek","Malinowski","Gorski","Wisniewska","Bednarek","Czerwinski","Cieslak","Witkowski","Andrzejewski",
    "Baranowski","Chmielewski","Cichon","Domanski","Frankowski","Gajewski","Halas","Idzik","Jaskolski","Kalinowski",
    "Lis","Majewski","Nawrocki","Olszewski","Perkowski","Rogalski","Sobczak","Szulc","Tarnowski","Urban",
    "Wasilewski","Zaremba","Zawisza","Bogdanowicz","Cwiklinski","Dobrzynski","Eliasz","Fabianski","Gluszek","Halasa"
)

$departments = @("IT","HR","Finance","Sales","Operations","Marketing","Support","Engineering")
$titlesByDepartment = @{
    "IT"         = @("Systems Administrator","Endpoint Engineer","Cloud Administrator","Helpdesk Specialist","Security Analyst")
    "HR"         = @("HR Specialist","Recruiter","HR Manager")
    "Finance"    = @("Accountant","Financial Analyst","Finance Manager")
    "Sales"      = @("Sales Representative","Account Executive","Sales Manager")
    "Operations" = @("Operations Analyst","Operations Coordinator","Operations Manager")
    "Marketing"  = @("Marketing Specialist","Content Coordinator","Marketing Manager")
    "Support"    = @("Support Engineer","Service Desk Analyst","Support Manager")
    "Engineering"= @("Engineer","Senior Engineer","Team Lead")
}

$offices = @(
    @{ City="Legnica"; State="Dolnoslaskie"; Country="PL"; CountryCode=616; Office="Legnica HQ"; Street="Zlotoryjska 12"; Postal="59-220" },
    @{ City="Wroclaw"; State="Dolnoslaskie"; Country="PL"; CountryCode=616; Office="Wroclaw"; Street="Powstancow Slaskich 5"; Postal="53-332" },
    @{ City="Krakow"; State="Malopolskie"; Country="PL"; CountryCode=616; Office="Krakow"; Street="Jasnogorska 1"; Postal="31-358" },
    @{ City="Poznan"; State="Wielkopolskie"; Country="PL"; CountryCode=616; Office="Poznan"; Street="Bukowska 18"; Postal="60-809" },
    @{ City="Gdansk"; State="Pomorskie"; Country="PL"; CountryCode=616; Office="Gdansk"; Street="Grunwaldzka 103"; Postal="80-244" }
)

$company = $CompanyName
$emailDomain = $effectiveUPNSuffix

# ----------------------------
# Groups
# ----------------------------
$secGroupPath = "OU=Security,$groupRoot"
$distGroupPath = "OU=Distribution,$groupRoot"
$appGroupPath = "OU=Applications,$groupRoot"
$licGroupPath = "OU=Licensing,$groupRoot"
$rbacGroupPath = "OU=RBAC,$groupRoot"

$groupsToCreate = @(
    @{ Name="GG-IT-Admins"; Path=$secGroupPath; Scope="Global"; Category="Security"; Description="IT admins lab group" },
    @{ Name="GG-HR-Users"; Path=$secGroupPath; Scope="Global"; Category="Security"; Description="HR users lab group" },
    @{ Name="GG-Finance-Users"; Path=$secGroupPath; Scope="Global"; Category="Security"; Description="Finance users lab group" },
    @{ Name="GG-Sales-Users"; Path=$secGroupPath; Scope="Global"; Category="Security"; Description="Sales users lab group" },
    @{ Name="GG-Workstations-LocalAdmins"; Path=$secGroupPath; Scope="Global"; Category="Security"; Description="Local admin test group" },
    @{ Name="GG-VPN-Users"; Path=$secGroupPath; Scope="Global"; Category="Security"; Description="VPN access test group" },
    @{ Name="GG-MFA-Excluded"; Path=$secGroupPath; Scope="Global"; Category="Security"; Description="Legacy exclusion test group" },
    @{ Name="GG-BreakGlass-Review"; Path=$secGroupPath; Scope="Global"; Category="Security"; Description="Review-only break glass test group" },
    @{ Name="DL-All-Staff"; Path=$distGroupPath; Scope="Universal"; Category="Distribution"; Description="All staff distribution group" },
    @{ Name="DL-Warsaw-Office"; Path=$distGroupPath; Scope="Universal"; Category="Distribution"; Description="Office distribution group" },
    @{ Name="DL-Krakow-Office"; Path=$distGroupPath; Scope="Universal"; Category="Distribution"; Description="Office distribution group" },
    @{ Name="APP-M365-E3-Licensed"; Path=$licGroupPath; Scope="Global"; Category="Security"; Description="License targeting test group" },
    @{ Name="APP-VISIO-Licensed"; Path=$licGroupPath; Scope="Global"; Category="Security"; Description="License targeting test group" },
    @{ Name="APP-PROJECT-Licensed"; Path=$licGroupPath; Scope="Global"; Category="Security"; Description="License targeting test group" },
    @{ Name="APP-VPN-Client"; Path=$appGroupPath; Scope="Global"; Category="Security"; Description="Application assignment test group" },
    @{ Name="APP-Adobe-Users"; Path=$appGroupPath; Scope="Global"; Category="Security"; Description="Application assignment test group" },
    @{ Name="RBAC-Helpdesk-L1"; Path=$rbacGroupPath; Scope="Global"; Category="Security"; Description="Helpdesk role testing" },
    @{ Name="RBAC-Helpdesk-L2"; Path=$rbacGroupPath; Scope="Global"; Category="Security"; Description="Helpdesk role testing" }
)

foreach ($g in $groupsToCreate) {
    Ensure-Group -Name $g.Name -Path $g.Path -Scope $g.Scope -Category $g.Category -Description $g.Description
}

# ----------------------------
# Generate manager seed users first
# ----------------------------
$managerUsers = @()
$managerDepartments = @("IT","HR","Finance","Sales","Operations","Marketing","Support","Engineering")

foreach ($dept in $managerDepartments) {
    $fn = Get-Random $firstNames
    $ln = Get-Random $lastNames
    $displayName = "$fn $ln"
    $baseSam = (($fn.Substring(0,1) + $ln).ToLower() -replace '[^a-z0-9]','')
    $userPath = "OU=$dept,$userRoot"
    $sam = Get-UniqueSam -BaseSam $baseSam -SearchBase $topOU
    $upn = "$sam@$emailDomain"
    $mail = "$sam@$emailDomain"
    $office = Get-Random $offices
    $title = "$dept Manager"

    if (-not (Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" -SearchBase $topOU -ErrorAction SilentlyContinue)) {
        New-ADUser `
            -Name $displayName `
            -GivenName $fn `
            -Surname $ln `
            -DisplayName $displayName `
            -SamAccountName $sam `
            -UserPrincipalName $upn `
            -EmailAddress $mail `
            -Path $userPath `
            -Enabled $true `
            -AccountPassword $securePassword `
            -ChangePasswordAtLogon $false `
            -Company $company `
            -Department $dept `
            -Title $title `
            -Office $office.Office `
            -City $office.City `
            -State $office.State `
            -Country $office.Country `
            -PostalCode $office.Postal `
            -StreetAddress $office.Street `
            -Description "Lab seeded manager account"

        Set-OptionalUserAttributes -Identity $sam -Replace @{
            # IDs / flags
            employeeID          = "MGR$((Get-Random -Minimum 1000 -Maximum 9999))"
            info                = $ext1      
            employeeType        = Get-Random @("A","B","C")              
            department          = $dept                     
            # Location
            co                  = "Poland"
            countryCode         = $office.CountryCode
            physicalDeliveryOfficeName = $office.Office

            # Mail & Entra
            proxyAddresses      = @("SMTP:$mail","smtp:$sam@$($dnsRoot)")
            #usageLocation       = "PL"  ###  Entra ID Attribute
        }

        $managerUsers += [PSCustomObject]@{
            Department = $dept
            Sam = $sam
            DistinguishedName = (Get-ADUser $sam).DistinguishedName
        }
        Write-Log "Created manager: $displayName ($dept)"
    }
}

# ----------------------------
# Generate standard users
# ----------------------------
$userCreated = 0
$specialMix = @("Standard","Standard","Standard","Standard","Pilot","Disabled","NonSync","Contractor","SharedStyle","Intern")

while ($userCreated -lt $UserCount) {
    $fn = Get-Random $firstNames
    $ln = Get-Random $lastNames
    $dept = Get-Random $departments
    $title = Get-Random $titlesByDepartment[$dept]
    $office = Get-Random $offices
    $scenario = Get-Random $specialMix

    $displayName = "$fn $ln"
    $baseSam = (($fn.Substring(0,1) + $ln + (Get-Random -Minimum 10 -Maximum 99)).ToLower() -replace '[^a-z0-9]','')
    $sam = Get-UniqueSam -BaseSam $baseSam -SearchBase $topOU
    $upn = "$sam@$emailDomain"
    $mail = "$sam@$emailDomain"

    switch ($scenario) {
        "Pilot" {
            $path = "OU=Users,$pilotRoot"
            $ext1 = "PILOT"
            $enabled = $true
            $description = "Pilot user for Entra Connect OU filtering"
        }
        "Disabled" {
            $path = "OU=Users,$disabledRoot"
            $ext1 = "DISABLED"
            $enabled = $false
            $description = "Disabled user for lifecycle and sync testing"
        }
        "NonSync" {
            $path = "OU=Users,$nonSyncRoot"
            $ext1 = "NOSYNC"
            $enabled = $true
            $description = "User intentionally placed outside sync scope"
        }
        "Contractor" {
            $path = "OU=Operations,$userRoot"
            $ext1 = "CONTRACTOR"
            $enabled = $true
            $description = "Contractor-style user"
        }
        "SharedStyle" {
            $path = "OU=HQ,$userRoot"
            $ext1 = "SHAREDSTYLE"
            $enabled = $true
            $description = "Shared-style user object for testing"
        }
        "Intern" {
            $path = "OU=Interns,$userRoot"
            $ext1 = "INTERN"
            $enabled = $true
            $description = "Intern user"
        }
        default {
            $path = "OU=$dept,$userRoot"
            $ext1 = "SYNC"
            $enabled = $true
            $description = "Standard synced lab user"
        }
    }

    if (-not ((Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" -SearchBase $topOU -ErrorAction SilentlyContinue) -and (Get-ADUser -LDAPFilter "(cn=$displayName)" -SearchBase $topOU -ErrorAction SilentlyContinue))) {
        New-ADUser `
            -Name $displayName `
            -GivenName $fn `
            -Surname $ln `
            -DisplayName $displayName `
            -SamAccountName $sam `
            -UserPrincipalName $upn `
            -EmailAddress $mail `
            -Path $path `
            -Enabled $enabled `
            -AccountPassword $securePassword `
            -ChangePasswordAtLogon $false `
            -PasswordNeverExpires:$false `
            -Company $company `
            -Department $dept `
            -Title $title `
            -Office $office.Office `
            -City $office.City `
            -State $office.State `
            -Country $office.Country `
            -PostalCode $office.Postal `
            -StreetAddress $office.Street `
            -Description $description `
            #-ErrorAction SilentlyContinue

        $mgr = $managerUsers | Where-Object Department -eq $dept | Get-Random
        $employeeId = "EMP$("{0:D5}" -f (Get-Random -Minimum 1 -Maximum 99999))"

        $proxy = @("SMTP:$mail","smtp:$sam@$dnsRoot")
        if ($scenario -eq "SharedStyle") {
            $proxy += "smtp:shared-$sam@$emailDomain"
        }

        $replace = @{
            employeeID          = $employeeId
            info                = $ext1                          
            department          = $dept                           
            l                   = $office.City                    
            description         = $(if ($enabled) { 'ACTIVE' } else { 'DISABLED' })
            employeeType        = Get-Random @("A","B","C")       
            co                  = "Poland"                        
            countryCode         = $office.CountryCode
            manager             = $mgr.DistinguishedName
            physicalDeliveryOfficeName = $office.Office
            proxyAddresses      = $proxy
            #usageLocation       = "PL"   ### Entra ID Attirbute
            #mailNickname        = $sam   ### Needs Exchange Schema
        }


        Set-OptionalUserAttributes -Identity $sam -Replace $replace

        if ($scenario -eq "Disabled") {
            Disable-ADAccount -Identity $sam
        }

        # Add to groups
        $groupAdds = @("DL-All-Staff","APP-M365-E3-Licensed")
        switch ($dept) {
            "IT"         { $groupAdds += @("GG-IT-Admins","GG-VPN-Users","RBAC-Helpdesk-L1","APP-VPN-Client") }
            "HR"         { $groupAdds += @("GG-HR-Users") }
            "Finance"    { $groupAdds += @("GG-Finance-Users") }
            "Sales"      { $groupAdds += @("GG-Sales-Users","APP-Adobe-Users") }
            "Support"    { $groupAdds += @("RBAC-Helpdesk-L1") }
        }
        if ($scenario -eq "Pilot") { $groupAdds += "GG-VPN-Users" }
        if ($scenario -eq "Contractor") { $groupAdds += "GG-MFA-Excluded" }
        if ((Get-Random -Minimum 1 -Maximum 100) -le 15) { $groupAdds += "APP-VISIO-Licensed" }
        if ((Get-Random -Minimum 1 -Maximum 100) -le 10) { $groupAdds += "APP-PROJECT-Licensed" }

        foreach ($groupName in ($groupAdds | Select-Object -Unique)) {
            try {
                Add-ADGroupMember -Identity $groupName -Members $sam -ErrorAction Stop
            } catch {}
        }

        $userCreated++
        Write-Log "Created user $userCreated/$UserCount : $displayName [$scenario]"
    }
}

# ----------------------------
# Service accounts / admin accounts / shared/resource style accounts
# ----------------------------
$specialAccounts = @(
    @{ Name="svc_aadconnect"; Sam="svc_aadconnect"; UPN="svc_aadconnect@$emailDomain"; Path="OU=Service Accounts,$topOU"; Desc="AAD Connect style service account"; Enabled=$true; Dept="IT"; Title="Service Account"; Ext1="SERVICE" },
    @{ Name="svc_sqlbackup"; Sam="svc_sqlbackup"; UPN="svc_sqlbackup@$emailDomain"; Path="OU=Service Accounts,$topOU"; Desc="SQL backup service account"; Enabled=$true; Dept="IT"; Title="Service Account"; Ext1="SERVICE" },
    @{ Name="svc_legacyapp"; Sam="svc_legacyapp"; UPN="svc_legacyapp@$emailDomain"; Path="OU=Service Accounts,$topOU"; Desc="Legacy app service account"; Enabled=$true; Dept="IT"; Title="Service Account"; Ext1="SERVICE" },
    @{ Name="adm.jnowak"; Sam="adm.jnowak"; UPN="adm.jnowak@$emailDomain"; Path="OU=Admins,$topOU"; Desc="Admin account"; Enabled=$true; Dept="IT"; Title="Privileged Admin"; Ext1="ADMIN" },
    @{ Name="adm.akowalska"; Sam="adm.akowalska"; UPN="adm.akowalska@$emailDomain"; Path="OU=Admins,$topOU"; Desc="Admin account"; Enabled=$true; Dept="IT"; Title="Privileged Admin"; Ext1="ADMIN" },
    @{ Name="shared.helpdesk"; Sam="shared.helpdesk"; UPN="shared.helpdesk@$emailDomain"; Path="OU=Shared Mailboxes,$resourceRoot"; Desc="Shared mailbox style object"; Enabled=$true; Dept="Support"; Title="Shared Resource"; Ext1="RESOURCE" },
    @{ Name="room.krk.101"; Sam="room.krk.101"; UPN="room.krk.101@$emailDomain"; Path="OU=Rooms,$resourceRoot"; Desc="Room mailbox style object"; Enabled=$true; Dept="Facilities"; Title="Room"; Ext1="RESOURCE" }
)

foreach ($acct in $specialAccounts) {
    if (-not (Get-ADUser -LDAPFilter "(sAMAccountName=$($acct.Sam))" -SearchBase $topOU -ErrorAction SilentlyContinue)) {
        $parts = $acct.Name -split '[\.\s]'
        $gn = $parts[0]
        $sn = if ($parts.Count -gt 1) { $parts[-1] } else { "Lab" }

        New-ADUser `
            -Name $acct.Name `
            -GivenName $gn `
            -Surname $sn `
            -DisplayName $acct.Name `
            -SamAccountName $acct.Sam `
            -UserPrincipalName $acct.UPN `
            -EmailAddress $acct.UPN `
            -Path $acct.Path `
            -Enabled $acct.Enabled `
            -AccountPassword $securePassword `
            -PasswordNeverExpires $true `
            -ChangePasswordAtLogon $false `
            -Company $company `
            -Department $acct.Dept `
            -Title $acct.Title `
            -Description $acct.Desc

        Set-OptionalUserAttributes -Identity $acct.Sam -Replace @{
            info                 = $acct.Ext1
            #usageLocation = "PL"  ### Entra ID Attribute
            proxyAddresses = @("SMTP:$($acct.UPN)")
            #mailNickname = $acct.Sam      ### Needs Exchange Schema
        }
        Write-Log "Created special account: $($acct.Name)"
    }
}

# ----------------------------
# Contacts
# ----------------------------
$contactPath = "OU=Contacts,$topOU"
for ($i = 1; $i -le 15; $i++) {
    $fn = Get-Random $firstNames
    $ln = Get-Random $lastNames
    $name = "$fn $ln External"
    $mail = ("external{0}@partner.example" -f $i)
    $existing = Get-ADObject -LDAPFilter "(&(objectClass=contact)(cn=$name))" -SearchBase $contactPath -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-ADObject -Name $name -Type contact -Path $contactPath -OtherAttributes @{
            displayName = $name
            mail = $mail
            #targetAddress = "SMTP:$mail"
            company = "PartnerCo"
            description = "External contact for GAL/sync lab"
        }
        Write-Log "Created contact: $name"
    }
}

# ----------------------------
# Computers
# ----------------------------
$computerTargets = @(
    @{ Prefix="PLWKS"; Path="OU=Workstations,$computerRoot"; OS="Windows 11 Enterprise"; Vers="23H2"; Count=[Math]::Floor($ComputerCount * 0.35) },
    @{ Prefix="PLLAP"; Path="OU=Laptops,$computerRoot"; OS="Windows 11 Enterprise"; Vers="23H2"; Count=[Math]::Floor($ComputerCount * 0.30) },
    @{ Prefix="PLKSK"; Path="OU=Kiosks,$computerRoot"; OS="Windows 10 Enterprise"; Vers="22H2"; Count=[Math]::Floor($ComputerCount * 0.10) },
    @{ Prefix="PLSRV"; Path="OU=Servers,$computerRoot"; OS="Windows Server 2022"; Vers="21H2"; Count=[Math]::Floor($ComputerCount * 0.15) },
    @{ Prefix="PLLEG"; Path="OU=Legacy,$computerRoot"; OS="Windows 10 Pro"; Vers="21H2"; Count=[Math]::Floor($ComputerCount * 0.05) },
    @{ Prefix="PLPIL"; Path="OU=Devices,$pilotRoot"; OS="Windows 11 Enterprise"; Vers="23H2"; Count=($ComputerCount - ([Math]::Floor($ComputerCount * 0.35) + [Math]::Floor($ComputerCount * 0.30) + [Math]::Floor($ComputerCount * 0.10) + [Math]::Floor($ComputerCount * 0.15) + [Math]::Floor($ComputerCount * 0.05))) }
)

foreach ($target in $computerTargets) {
    for ($i = 1; $i -le $target.Count; $i++) {
        $name = "{0}{1:D3}" -f $target.Prefix, $i
        if (-not (Get-ADComputer -LDAPFilter "(cn=$name)" -SearchBase $topOU -ErrorAction SilentlyContinue)) {
            New-ADComputer `
                -Name $name `
                -SamAccountName $name `
                -Path $target.Path `
                -Enabled $true `
                -OperatingSystem $target.OS `
                -OperatingSystemVersion $target.Vers `
                -Description "Lab seeded computer"
            try {
                Set-ADComputer -Identity $name -Add @{
                    location = (Get-Random $offices).Office
                }
            } catch {}
            Write-Log "Created computer: $name"
        }
    }
}

# Add a few disabled and non-sync computers
$extraComputers = @(
    @{ Name="PLDIS001"; Path="OU=Computers,$disabledRoot"; Enabled=$false; OS="Windows 10 Enterprise"; Vers="22H2"; Desc="Disabled device object" },
    @{ Name="PLDIS002"; Path="OU=Computers,$disabledRoot"; Enabled=$false; OS="Windows 11 Enterprise"; Vers="23H2"; Desc="Disabled device object" },
    @{ Name="PLNS001"; Path="OU=Computers,$nonSyncRoot"; Enabled=$true; OS="Windows 10 Enterprise"; Vers="22H2"; Desc="Non-sync test device object" },
    @{ Name="PLNS002"; Path="OU=Computers,$nonSyncRoot"; Enabled=$true; OS="Windows 11 Enterprise"; Vers="23H2"; Desc="Non-sync test device object" }
)

foreach ($c in $extraComputers) {
    if (-not (Get-ADComputer -LDAPFilter "(cn=$($c.Name))" -SearchBase $topOU -ErrorAction SilentlyContinue)) {
        New-ADComputer -Name $c.Name -SamAccountName $c.Name -Path $c.Path -Enabled $c.Enabled -OperatingSystem $c.OS -OperatingSystemVersion $c.Vers -Description $c.Desc
        Write-Log "Created extra computer: $($c.Name)"
    }
}

# ----------------------------
# Nested group examples
# ----------------------------
$nestedExamples = @(
    @{ Parent="GG-IT-Admins"; Child="RBAC-Helpdesk-L2" },
    @{ Parent="APP-M365-E3-Licensed"; Child="GG-HR-Users" },
    @{ Parent="APP-M365-E3-Licensed"; Child="GG-Finance-Users" },
    @{ Parent="APP-M365-E3-Licensed"; Child="GG-Sales-Users" }
)
foreach ($n in $nestedExamples) {
    try { Add-ADGroupMember -Identity $n.Parent -Members $n.Child -ErrorAction Stop } catch {}
}

# ----------------------------
# Summary
# ----------------------------
$userTotal = (Get-ADUser -SearchBase $topOU -LDAPFilter "(objectClass=user)").Count
$computerTotal = (Get-ADComputer -SearchBase $topOU -LDAPFilter "(objectClass=computer)").Count
$groupTotal = (Get-ADGroup -SearchBase $topOU -LDAPFilter "(objectClass=group)").Count
$contactTotal = (Get-ADObject -SearchBase $topOU -LDAPFilter "(objectClass=contact)").Count
$ouTotal = (Get-ADOrganizationalUnit -SearchBase $topOU -LDAPFilter "(objectClass=organizationalUnit)").Count

Write-Host ""
Write-Host "========== LAB BUILD COMPLETE ==========" -ForegroundColor Green
Write-Host "Top OU        : LAB"
Write-Host "Users         : $userTotal"
Write-Host "Computers     : $computerTotal"
Write-Host "Groups        : $groupTotal"
Write-Host "Contacts      : $contactTotal"
Write-Host "OUs           : $ouTotal"
Write-Host "Default pass  : $DefaultPassword"
Write-Host "UPN suffix    : $effectiveUPNSuffix"
Write-Host ""
Write-Host "Suggested Entra Connect sync scope tests:"
Write-Host "1. Sync only OU=EntraPilot,OU=LAB,..."
Write-Host "2. Sync OU=Users and OU=Computers but exclude OU=NonSync"
Write-Host "3. Test disabled accounts in OU=Disabled Objects"
Write-Host "4. Test group sync and nested groups"
Write-Host "5. Test device/computer object visibility and OU filtering"
Write-Host "========================================"