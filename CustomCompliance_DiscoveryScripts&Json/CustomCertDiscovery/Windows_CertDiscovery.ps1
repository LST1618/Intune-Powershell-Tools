$storePath = 'Cert:\LocalMachine\My'
$subject = 'CN=WhateverDomain Test Certificate'
$issuer = 'CN=WhateverDomain Test Certificate'
$minimumDaysRemaining = 30

$matchingCert = Get-ChildItem -Path $storePath -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Subject -like "*$subject*" -and
        $_.Issuer -like "*$issuer*"
    } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

$certPresent = $null -ne $matchingCert
$notExpired = $false
$daysRemaining = -1

if ($certPresent) {
    $daysRemaining = [int](New-TimeSpan -Start (Get-Date) -End $matchingCert.NotAfter).TotalDays
    $notExpired = $daysRemaining -ge $minimumDaysRemaining
}

$result = @{
    CertPresent = $certPresent
    CertNotExpired30Days = $notExpired
    CertThumbprint = if ($certPresent) { $matchingCert.Thumbprint } else { "" }
    CertSubject = if ($certPresent) { $matchingCert.Subject } else { "" }
    CertIssuer = if ($certPresent) { $matchingCert.Issuer } else { "" }
}

return $result | ConvertTo-Json