$av = Get-CimInstance `
    -Namespace 'root\SecurityCenter2' `
    -ClassName AntiVirusProduct

$expected = $av | Where-Object {
    $_.displayName -match 'Sophos'
}

if ($expected) {
    $result = @{
    AntivirusRegistered = $true 
    } 
} else {
    $result = @{
    AntivirusRegistered = $false 
    } 
}

return $result | ConvertTo-Json