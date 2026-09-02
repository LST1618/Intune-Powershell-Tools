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


$cert = New-SelfSignedCertificate `
    -Subject 'CN=WhateverDomain Test Certificate' `
    -CertStoreLocation 'Cert:\LocalMachine\My' `
    -Type Custom `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddDays(15)
	
	
Export-Certificate `
    -Cert $cert `
    -FilePath "$env:USERPROFILE\Desktop\WhateverDomainTestCertificate.cer"




Install-Module Office365DnsChecker
Import-Module Office365DnsChecker
Test-Office365DnsRecords whateverdomain1618.online


$domain = "whateverdomain1618.online"

"NS"
Resolve-DnsName $domain -Type NS | Select Name, NameHost

"MX"
Resolve-DnsName $domain -Type MX | Sort Preference | Select Name, NameExchange, Preference

"TXT"
Resolve-DnsName $domain -Type TXT | Select -ExpandProperty Strings

"Autodiscover"
Resolve-DnsName "autodiscover.$domain" -Type CNAME | Select Name, NameHost

"SPF"
Resolve-DnsName $domain -Type TXT | Where-Object { $_.Strings -match "v=spf1" } | Select -ExpandProperty Strings


### AD Automation
Install-WindowsFeature -Name RSAT-AD-Tools -IncludeAllSubFeature -verbose

Import-Module ActiveDirectory -verbose


##Full Hostname / FQDN
[System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
#PS C:\Users\Administrator> [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
#DC1.corp.contoso.com