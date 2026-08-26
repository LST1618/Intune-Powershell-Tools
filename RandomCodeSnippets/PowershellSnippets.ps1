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