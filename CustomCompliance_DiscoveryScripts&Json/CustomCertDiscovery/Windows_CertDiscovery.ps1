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
}

return $result | ConvertTo-Json -Compress