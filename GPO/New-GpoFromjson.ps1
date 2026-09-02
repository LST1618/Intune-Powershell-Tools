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

#.\New-GpoSeed.ps1 -ConfigPath .\sample-gpo-seed.json -LogFolder .\Logs

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath,

    [string]$LogFolder = ".\\Logs",

    [switch]$SkipTranscript
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RunId = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:CsvLog = $null
$script:TranscriptPath = $null

function Initialize-Logging {
    param([string]$Folder)

    if (-not (Test-Path -LiteralPath $Folder)) {
        New-Item -Path $Folder -ItemType Directory -Force | Out-Null
    }

    $script:CsvLog = Join-Path $Folder ("GpoSeed_{0}.csv" -f $script:RunId)
    $script:TranscriptPath = Join-Path $Folder ("GpoSeed_{0}.log" -f $script:RunId)

    "Timestamp,Level,Phase,GPO,Action,Target,Result,Message" | Out-File -FilePath $script:CsvLog -Encoding utf8 -Force

    if (-not $SkipTranscript) {
        Start-Transcript -Path $script:TranscriptPath -Force | Out-Null
    }
}

function Stop-Logging {
    if (-not $SkipTranscript) {
        try { Stop-Transcript | Out-Null } catch { }
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [string]$Phase = 'General',
        [string]$Gpo = '',
        [string]$Action = '',
        [string]$Target = '',
        [string]$Result = ''
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] [$Level] [$Phase] $Message"

    $row = [pscustomobject]@{
        Timestamp = $timestamp
        Level     = $Level
        Phase     = $Phase
        GPO       = $Gpo
        Action    = $Action
        Target    = $Target
        Result    = $Result
        Message   = $Message
    }
    $row | Export-Csv -Path $script:CsvLog -NoTypeInformation -Append -Encoding utf8
}

function Assert-Module {
    if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
        throw 'GroupPolicy module not found. Install GPMC / RSAT Group Policy Management tools.'
    }
    Import-Module GroupPolicy -ErrorAction Stop
}




function Test-DomainDn {
    param([string]$Dn)
    return ($Dn -match '^((DC|OU|CN)=[^,]+,?)+$')
}

function Validate-Enum {
    param(
        [string]$Value,
        [string[]]$Allowed,
        [string]$FieldName
    )
    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "$FieldName cannot be empty. Allowed: $($Allowed -join ', ')"
    }
    if ($Allowed -notcontains $Value) {
        throw "$FieldName '$Value' is invalid. Allowed: $($Allowed -join ', ')"
    }
}

function Validate-Config {
    param([object]$Config)

    if ([string]::IsNullOrWhiteSpace($Config.DomainFqdn)) {
        throw 'DomainFqdn is required.'
    }
    if ([string]::IsNullOrWhiteSpace($Config.Server)) {
        throw 'Server is required.'
    }
    if (-not $Config.Gpos -or $Config.Gpos.Count -lt 1) {
        throw 'At least one GPO definition is required in Gpos.'
    }

    $seen = @{}
    foreach ($gpo in $Config.Gpos) {
        if ([string]::IsNullOrWhiteSpace($gpo.Name)) {
            throw 'Each GPO must have a non-empty Name.'
        }
        if ($seen.ContainsKey($gpo.Name)) {
            throw "Duplicate GPO name detected: $($gpo.Name)"
        }
        $seen[$gpo.Name] = $true

        if ($gpo.Status) {
            Validate-Enum -Value $gpo.Status -Allowed @('AllSettingsEnabled','UserSettingsDisabled','ComputerSettingsDisabled','AllSettingsDisabled') -FieldName "GPO Status for '$($gpo.Name)'"
        }



        foreach ($reg in @($gpo.RegistryPolicies))  {
 
                    
            Validate-Enum -Value $reg.Context -Allowed @('Computer','User') -FieldName "RegistryPolicies.Context for '$($gpo.Name)'"
            Validate-Enum -Value $reg.Type -Allowed @('String','ExpandString','Binary','DWord','MultiString','QWord') -FieldName "RegistryPolicies.Type for '$($gpo.Name)'"
            if ([string]::IsNullOrWhiteSpace($reg.Key) -or [string]::IsNullOrWhiteSpace($reg.ValueName)) {
                throw "RegistryPolicies entry for '$($gpo.Name)' must include Key and ValueName."
            }
            if ($reg.Context -eq 'Computer' -and $reg.Key -notmatch '^HKLM\\') {
                throw "Computer policy in '$($gpo.Name)' must use HKLM\\ path: $($reg.Key)"
            }
            if ($reg.Context -eq 'User' -and $reg.Key -notmatch '^HKCU\\') {
                throw "User policy in '$($gpo.Name)' must use HKCU\\ path: $($reg.Key)"
            }
        }

        foreach ($sf in @($gpo.SecurityFiltering)) {
            Validate-Enum -Value $sf.TargetType -Allowed @('Group','User','Computer') -FieldName "SecurityFiltering.TargetType for '$($gpo.Name)'"
            Validate-Enum -Value $sf.PermissionLevel -Allowed @('GpoRead','GpoApply','GpoEdit','GpoEditDeleteModifySecurity','None') -FieldName "SecurityFiltering.PermissionLevel for '$($gpo.Name)'"
            if ([string]::IsNullOrWhiteSpace($sf.TargetName)) {
                throw "SecurityFiltering.TargetName is required for '$($gpo.Name)'."
            }
        }

        foreach ($link in @($gpo.Links)) {
            if ([string]::IsNullOrWhiteSpace($link.Target)) {
                throw "Links.Target is required for '$($gpo.Name)'."
            }
            if (-not (Test-DomainDn -Dn $link.Target)) {
                throw "Links.Target does not look like a distinguished name for '$($gpo.Name)': $($link.Target)"
            }
            if ($null -ne $link.Order -and [int]$link.Order -lt 1) {
                throw "Links.Order must be 1 or higher for '$($gpo.Name)'."
            }
        }

        if ($gpo.WmiFilter) {
            if ([string]::IsNullOrWhiteSpace($gpo.WmiFilter.Name)) {
                throw "WmiFilter.Name is required for '$($gpo.Name)'."
            }
            if ($Config.CreateMissingWmiFilters -and [string]::IsNullOrWhiteSpace($gpo.WmiFilter.Query)) {
                throw "WmiFilter.Query is required when CreateMissingWmiFilters is true for '$($gpo.Name)'."
            }
        }
    }
}

function Test-AdObjectExists {
    param(
        [string]$DistinguishedName,
        [string]$Server
    )
    try {
        $path = "LDAP://$Server/$DistinguishedName"
        $obj = [ADSI]$path
        return ($null -ne $obj.distinguishedName)
    }
    catch {
        return $false
    }
}

function Get-GpmEnvironment {
    $gpm = New-Object -ComObject GPMgmt.GPM
    $constants = $gpm.GetConstants()
    return [pscustomobject]@{
        Gpm = $gpm
        Constants = $constants
    }
}

function Get-GpmDomain {
    param(
        [object]$Gpm,
        [object]$Constants,
        [string]$Domain,
        [string]$Server
    )
    return $Gpm.GetDomain($Domain, $Server, $Constants.UseAnyDC)
}

function Ensure-Gpo {
    param(
        [string]$Name,
        [string]$Comment,
        [string]$Status,
        [string]$Domain,
        [string]$Server
    )

    $gpo = Get-GPO -Name $Name -Domain $Domain -Server $Server -ErrorAction SilentlyContinue
    if (-not $gpo) {
        Write-Log -Message "Creating GPO '$Name'" -Phase 'Create' -Gpo $Name -Action 'New-GPO' -Result 'Started'
        $newParams = @{ Name = $Name; Domain = $Domain; Server = $Server }
        if ($Comment) { $newParams.Comment = $Comment }
        $gpo = New-GPO @newParams
        Write-Log -Message "Created GPO '$Name'" -Phase 'Create' -Gpo $Name -Action 'New-GPO' -Result 'Success'
    }
    
    return (Get-GPO -Name $Name -Domain $Domain -Server $Server)
}

function Set-RegistryPolicies {
    param(
        [string]$GpoName,
        [array]$Policies,
        [string]$Domain,
        [string]$Server
    )

    foreach ($policy in @($Policies)) {
        $params = @{
            Name      = $GpoName
            Key       = $policy.Key
            ValueName = $policy.ValueName
            Type      = $policy.Type
            Domain    = $Domain
            Server    = $Server
        }
        if ($null -ne $policy.Value) { $params.Value = $policy.Value }

        Set-GPRegistryValue @params | Out-Null
        Write-Log -Message "Applied registry policy '$($policy.Key)\\$($policy.ValueName)'" -Phase 'Registry' -Gpo $GpoName -Action 'Set-GPRegistryValue' -Target $policy.Key -Result 'Success'
    }
}

function Set-SecurityFiltering {
    param(
        [string]$GpoName,
        [array]$Entries,
        [string]$Domain,
        [string]$Server,
        [bool]$RemoveAuthenticatedUsersApply
    )

    if ($RemoveAuthenticatedUsersApply) {
        Set-GPPermission -Name $GpoName -TargetName 'Authenticated Users' -TargetType Group -PermissionLevel None -Replace -Confirm:$false -DomainName $Domain -Server $Server | Out-Null
        Write-Log -Message "Removed Authenticated Users apply permission" -Phase 'Security' -Gpo $GpoName -Action 'Set-GPPermission' -Target 'Authenticated Users' -Result 'Success'
    }

    foreach ($entry in @($Entries)) {
        Set-GPPermission -Name $GpoName -TargetName $entry.TargetName -TargetType $entry.TargetType -PermissionLevel $entry.PermissionLevel -Replace -Confirm:$false -DomainName $Domain -Server $Server | Out-Null
        Write-Log -Message "Applied security filtering for '$($entry.TargetName)'" -Phase 'Security' -Gpo $GpoName -Action 'Set-GPPermission' -Target $entry.TargetName -Result 'Success'
    }
}

function Get-OrCreateWmiFilter {
    param(
        [object]$GpmDomain,
        [object]$Constants,
        [string]$Name,
        [string]$Description,
        [string]$Query,
        [bool]$CreateIfMissing
    )

    $criteria = $GpmDomain.CreateSearchCriteria()
    $criteria.Add($Constants.SearchPropertyWMIFilterName, $Constants.SearchOpEquals, $Name)
    $filters = $GpmDomain.SearchWMIFilters($criteria)

    if ($filters.Count -gt 0) {
        return $filters.Item(1)
    }

    if (-not $CreateIfMissing) {
        throw "WMI filter '$Name' not found and creation is disabled."
    }

    $filter = $GpmDomain.CreateWMIFilter()
    $filter.Name = $Name
    if ($Description) { $filter.Description = $Description }
    $filter.AddQuery('WQL', $Query)
    $filter.Save()
    return $filter
}

function Set-WmiFilterAssignment {
    param(
        [object]$GpmDomain,
        [object]$Constants,
        [string]$GpoName,
        [object]$WmiFilter
    )

    $criteria = $GpmDomain.CreateSearchCriteria()
    $criteria.Add($Constants.SearchPropertyGPODisplayName, $Constants.SearchOpEquals, $GpoName)
    $gpos = $GpmDomain.SearchGPOs($criteria)
    if ($gpos.Count -lt 1) {
        throw "Unable to locate GPO '$GpoName' in GPMC COM search."
    }

    $gpo = $gpos.Item(1)
    $gpo.SetWMIFilter($WmiFilter)
    $gpo.Save()
    Write-Log -Message "Assigned WMI filter '$($WmiFilter.Name)'" -Phase 'WMI' -Gpo $GpoName -Action 'SetWMIFilter' -Target $WmiFilter.Name -Result 'Success'
}

function Ensure-Links {
    param(
        [Guid]$GpoGuid,
        [string]$GpoName,
        [array]$Links,
        [string]$Domain,
        [string]$Server
    )

    foreach ($link in @($Links)) {
        $enabled = if ($link.Enabled -eq $false) { 'No' } else { 'Yes' }
        $enforced = if ($link.Enforced -eq $true) { 'Yes' } else { 'No' }
        $order = if ($null -ne $link.Order) { [int]$link.Order } else { 1 }

        if (-not (Test-AdObjectExists -DistinguishedName $link.Target -Server $Server)) {
            throw "Link target not found in AD: $($link.Target)"
        }

        try {
            New-GPLink -Guid $GpoGuid -Target $link.Target -LinkEnabled $enabled -Enforced $enforced -Order $order -Domain $Domain -Server $Server | Out-Null
            Write-Log -Message "Created link to '$($link.Target)'" -Phase 'Links' -Gpo $GpoName -Action 'New-GPLink' -Target $link.Target -Result 'Success'
        }
        catch {
            Set-GPLink -Guid $GpoGuid -Target $link.Target -LinkEnabled $enabled -Enforced $enforced -Order $order -Domain $Domain -Server $Server | Out-Null
            Write-Log -Message "Updated existing link to '$($link.Target)'" -Level 'WARN' -Phase 'Links' -Gpo $GpoName -Action 'Set-GPLink' -Target $link.Target -Result 'Updated'
        }
    }
}

try {
    Initialize-Logging -Folder $LogFolder
    Write-Log -Message 'Starting GPO seed run.' -Phase 'Init' -Action 'Initialize' -Result 'Success'

    Assert-Module
    $config = Get-Content -LiteralPath $ConfigPath | ConvertFrom-Json
    Validate-Config -Config $config

    $gpmEnv = Get-GpmEnvironment
    $gpmDomain = Get-GpmDomain -Gpm $gpmEnv.Gpm -Constants $gpmEnv.Constants -Domain $config.DomainFqdn -Server $config.Server

    foreach ($gpoDef in @($config.Gpos)) {
        $gpo = Ensure-Gpo -Name $gpoDef.Name -Comment $gpoDef.Comment -Status $gpoDef.Status -Domain $config.DomainFqdn -Server $config.Server

        if ($gpoDef.RegistryPolicies) {
            Set-RegistryPolicies -GpoName $gpo.DisplayName -Policies $gpoDef.RegistryPolicies -Domain $config.DomainFqdn -Server $config.Server
        }

        if ($gpoDef.SecurityFiltering) {
            Set-SecurityFiltering -GpoName $gpo.DisplayName -Entries $gpoDef.SecurityFiltering -Domain $config.DomainFqdn -Server $config.Server -RemoveAuthenticatedUsersApply ([bool]$config.RemoveAuthenticatedUsersApply)
        }

        if ($gpoDef.WmiFilter) {
            $wmi = Get-OrCreateWmiFilter -GpmDomain $gpmDomain -Constants $gpmEnv.Constants -Name $gpoDef.WmiFilter.Name -Description $gpoDef.WmiFilter.Description -Query $gpoDef.WmiFilter.Query -CreateIfMissing ([bool]$config.CreateMissingWmiFilters)
            Set-WmiFilterAssignment -GpmDomain $gpmDomain -Constants $gpmEnv.Constants -GpoName $gpo.DisplayName -WmiFilter $wmi
        }

        if ($gpoDef.Links) {
            Ensure-Links -GpoGuid $gpo.Id -GpoName $gpo.DisplayName -Links $gpoDef.Links -Domain $config.DomainFqdn -Server $config.Server
        }
    }

    Write-Log -Message 'GPO seed run completed successfully.' -Phase 'Done' -Action 'Complete' -Result 'Success'
}
catch {
    Write-Log -Message $_.Exception.Message -Level 'ERROR' -Phase 'Fatal' -Action 'Abort' -Result 'Failed'
    throw
}
finally {
    Stop-Logging
}
