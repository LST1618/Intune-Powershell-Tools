<#
.SYNOPSIS
    Comprehensive Intune Index Analysis 2.0 — FULLY READ-ONLY

.NOTES
    Version     : 2.0
    Requires    : PowerShell 7+
    Modules     : Microsoft.Graph, ImportExcel
    Permissions : Group.Read.All, GroupMember.Read.All,
                DeviceManagementConfiguration.Read.All



MIT License

Copyright (c) 2026 

Łukasz Stachów
https://github.com/LST1618
https://www.linkedin.com/in/%C5%82ukasz-stach%C3%B3w-b1b886270/

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [string]$OutputPath        = "",
    [int]$DeepNestingThreshold = 3
)


if (-not $OutputPath) {
    $OutputPath = "C:\Temp\IntuneIndexAnalysis2.0_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
}

$RequiredScopes = @(
    "Group.Read.All",
    "GroupMember.Read.All",
    "DeviceManagementConfiguration.Read.All"
)

#region ══════════════════════════════════════════════════════
#  HELPERS
#══════════════════════════════════════════════════════════════

function Write-Log {
    param(
        [string]$Msg,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR","SECTION")][string]$Level = "INFO"
    )
    $ts     = Get-Date -Format "HH:mm:ss"
    $color  = @{ INFO="White"; SUCCESS="Green"; WARN="Yellow"; ERROR="Red"; SECTION="Cyan" }[$Level]
    $prefix = @{ INFO="  .  "; SUCCESS="  OK "; WARN="  !! "; ERROR="  XX "; SECTION="===" }[$Level]
    Write-Host "[$ts]$prefix $Msg" -ForegroundColor $color
}

function Get-OdataType {
    param([object]$Obj)
    if ($null -eq $Obj)                                        { return $null }
    if ($Obj -is [hashtable])                                  { return $Obj['@odata.type'] }
    if ($Obj -is [System.Management.Automation.PSObject] -and
        $Obj.PSObject.Properties['@odata.type'])               { return $Obj.'@odata.type' }
    return $null
}

function Get-ObjId {
    param([object]$Obj)
    if ($null -eq $Obj)              { return $null }
    if ($Obj -is [hashtable])        { return $Obj['id'] }
    if ($Obj -is [System.Management.Automation.PSObject] -and
        $Obj.PSObject.Properties['id']) { return $Obj.id }
    return $null
}

function Get-ObjDisplayName {
    param([object]$Obj)
    if ($null -eq $Obj)              { return $null }
    if ($Obj -is [hashtable])        { return $Obj['displayName'] }
    if ($Obj -is [System.Management.Automation.PSObject] -and
        $Obj.PSObject.Properties['displayName']) { return $Obj.displayName }
    return $null
}

function Get-AnyProp {
    param([object]$Obj, [string[]]$Names)
    if ($null -eq $Obj) { return $null }
    foreach ($n in $Names) {
        if ($Obj -is [hashtable] -and $Obj.ContainsKey($n))    { return $Obj[$n] }
        if ($Obj -is [System.Management.Automation.PSObject] -and
            $Obj.PSObject.Properties[$n])                       { return $Obj.$n }
    }
    return $null
}

function Invoke-GraphGetAll {
    param([string]$Uri, [switch]$Silent)
    $results = [System.Collections.Generic.List[object]]::new()
    do {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
            if ($resp.PSObject.Properties['value'] -and $resp.value) {
                $results.AddRange([object[]]$resp.value)
            }
            $Uri = if ($resp.PSObject.Properties['@odata.nextLink']) {
                        $resp.'@odata.nextLink'
                    } else { $null }
        } catch {
            if (-not $Silent) { Write-Log "Graph error $Uri : $($_.Exception.Message)" WARN }
            break
        }
    } while ($Uri)

    return $results
}

function Invoke-GraphBatch {
    param([string[]]$GroupIds, [string]$MemberEndpoint)

    $map       = @{}
    $batchSize = 20

    for ($i = 0; $i -lt $GroupIds.Count; $i += $batchSize) {
        $chunk = @($GroupIds[$i..([Math]::Min($i + $batchSize - 1, $GroupIds.Count - 1))])

        $requests = [System.Collections.Generic.List[object]]::new()
        for ($ri = 0; $ri -lt $chunk.Count; $ri++) {

            $requests.Add([PSCustomObject]@{
                id     = "$($ri + 1)|||$($chunk[$ri])"
                method = "GET"
                url    = "/groups/$($chunk[$ri])/$($MemberEndpoint)?`$select=id,displayName&`$top=999"
            })
        }

        $body = [PSCustomObject]@{ requests = $requests.ToArray() } |
                ConvertTo-Json -Depth 4 -Compress

        try {
            $batchResp = Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/`$batch" `
                -Body $body -ContentType "application/json" -OutputType PSObject

            foreach ($r in $batchResp.responses) {
                $gid = ($r.id -split '\|\|\|', 2)[1]
                if ($r.status -eq 200) {
                    if ($null -ne $r.body -and
                        $r.body.PSObject.Properties['value'] -and
                        $null -ne $r.body.value) {
                        $map[$gid] = @($r.body.value)
                    } else {
                        $map[$gid] = @()
                    }
                } else {
                    $map[$gid] = @()
                    Write-Log "Batch sub-request HTTP $($r.status) for group $gid" WARN
                }
            }
        } catch {
            Write-Log "Batch POST failed for chunk at index $i — falling back to individual GETs..." WARN
            foreach ($gid in $chunk) {
                try {
                    $map[$gid] = @(Invoke-GraphGetAll `
                        -Uri "https://graph.microsoft.com/v1.0/groups/$gid/$($MemberEndpoint)?`$select=id,displayName" `
                        -Silent)
                } catch { $map[$gid] = @() }
            }
        }
        Start-Sleep -Milliseconds 150
    }
    return $map
}

function Build-NestingTree {
    param(
        [string]$GroupId,
        [string]$RootId,
        [int]$Level,
        [int]$NestThreshold,
        [hashtable]$DirectMap,
        [hashtable]$GroupLookup,
        [System.Collections.Generic.HashSet[string]]$Visited,
        [string]$Path,
        [string]$PathIds
    )

    $rows = [System.Collections.Generic.List[object]]::new()

    if ($Level -ge 20) {
        $rows.Add([PSCustomObject]@{
            RootGroupId     = $RootId
            RootGroupName   = $GroupLookup[$RootId]  ?? $RootId
            ParentGroupId   = $GroupId
            ParentGroupName = $GroupLookup[$GroupId] ?? $GroupId
            ChildGroupId    = "DEPTH-LIMIT"
            ChildGroupName  = "MAX DEPTH 20 REACHED"
            NestingDepth    = $Level
            PathChain       = "$Path -> [MAX DEPTH 20]"
            PathChainIds    = $PathIds
            IsCircular      = $false
            RiskLevel       = "CRITICAL"
            RiskReason      = "Nesting depth hit hard limit of 20 — investigate for near-circular structure"
            Recommendation  = "INVESTIGATE: Chain exceeds 20 levels. Map manually and flatten aggressively."
        })
        return ,$rows  # comma operator prevents single-item list unwrap
    }

    if ($Visited.Contains($GroupId)) {
        $rows.Add([PSCustomObject]@{
            RootGroupId     = $RootId
            RootGroupName   = $GroupLookup[$RootId]  ?? $RootId
            ParentGroupId   = $GroupId
            ParentGroupName = $GroupLookup[$GroupId] ?? $GroupId
            ChildGroupId    = $GroupId
            ChildGroupName  = $GroupLookup[$GroupId] ?? $GroupId
            NestingDepth    = $Level
            PathChain       = "$Path -> *** CIRCULAR REFERENCE ***"
            PathChainIds    = $PathIds
            IsCircular      = $true
            RiskLevel       = "CRITICAL"
            RiskReason      = "Circular reference — group contains itself or an ancestor"
            Recommendation  = "URGENT: Open Entra portal > Groups > '$($GroupLookup[$RootId] ?? $RootId)' > Members. Remove the child that creates the loop back to an ancestor."
        })
        return ,$rows # comma operator prevents single-item list unwrap
    }

    $Visited.Add($GroupId) | Out-Null

    $directMembers = if ($DirectMap.ContainsKey($GroupId)) { @($DirectMap[$GroupId]) } else { @() }
    $childGroups   = @($directMembers | Where-Object { (Get-OdataType $_) -eq '#microsoft.graph.group' })

    foreach ($child in $childGroups) {
        $childId    = Get-ObjId $child
        $childName  = $GroupLookup[$childId]  ?? $childId
        $parentName = $GroupLookup[$GroupId]  ?? $GroupId
        $rootName   = $GroupLookup[$RootId]   ?? $RootId
        $newPath    = "$Path -> $childName"
        $newPathIds = "$PathIds|$childId"

        $riskLevel = if     ($Level + 1 -ge $NestThreshold + 2) { "CRITICAL" }
                    elseif ($Level + 1 -ge $NestThreshold)      { "HIGH" }
                    elseif ($Level + 1 -eq 2)                   { "MEDIUM" }
                    else                                         { "LOW" }

        $riskReason = switch ($riskLevel) {
            "CRITICAL" { "Extreme nesting depth ($($Level+1) levels) — serious Intune policy evaluation overhead" }
            "HIGH"     { "Deep nesting ($($Level+1) levels) — exceeds recommended depth of $($NestThreshold - 1)" }
            "MEDIUM"   { "Double nesting (2 levels) — review if intentional" }
            "LOW"      { "Single nesting (1 level) — generally acceptable" }
        }

        $recommendation = switch ($riskLevel) {
            "CRITICAL" { "FLATTEN IMMEDIATELY: At $($Level+1) levels Intune walks the full chain per policy check. Add members of nested groups directly to '$rootName' and remove intermediate layers." }
            "HIGH"     { "REVIEW AND FLATTEN: Depth $($Level+1) exceeds threshold. Add members of '$childName' directly to '$rootName' or assign the policy directly to '$childName'." }
            "MEDIUM"   { "REVIEW: Check if intermediate group '$parentName' is necessary. If not, move '$childName' members directly into '$rootName'." }
            "LOW"      { "MONITOR: Single-level nesting is acceptable. Verify '$childName' membership is intentional and document the reason in the group description." }
        }

        $rows.Add([PSCustomObject]@{
            RootGroupId     = $RootId
            RootGroupName   = $rootName
            ParentGroupId   = $GroupId
            ParentGroupName = $parentName
            ChildGroupId    = $childId
            ChildGroupName  = $childName
            NestingDepth    = $Level + 1
            PathChain       = $newPath
            PathChainIds    = $newPathIds
            IsCircular      = $false
            RiskLevel       = $riskLevel
            RiskReason      = $riskReason
            Recommendation  = $recommendation
        })

        $visitedClone = [System.Collections.Generic.HashSet[string]]::new($Visited)


        $childRows = @(Build-NestingTree `
            -GroupId       $childId `
            -RootId        $RootId `
            -Level         ($Level + 1) `
            -NestThreshold $NestThreshold `
            -DirectMap     $DirectMap `
            -GroupLookup   $GroupLookup `
            -Visited       $visitedClone `
            -Path          $newPath `
            -PathIds       $newPathIds)
        if ($childRows.Count -gt 0) { $rows.AddRange([object[]]$childRows) }
    }

    return ,$rows  # comma operator on the normal return path too
}

function Get-IntuneAssignmentIndex {
    Write-Log "Building Intune assignment index (22 policy types)..." INFO
    $assignIdx = @{}

    $eps = @(
        @{ L="Device Config";         U="https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Settings Catalog";      U="https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$expand=assignments&`$select=id,name,assignments" }
        @{ L="Admin Templates";       U="https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Compliance Policy";     U="https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Win32 App";             U="https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')&`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="App Config Policy";     U="https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="PowerShell Script";     U="https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Proactive Remediation"; U="https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Feature Update";        U="https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Driver Update";         U="https://graph.microsoft.com/beta/deviceManagement/windowsDriverUpdateProfiles?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Enrollment Config";     U="https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="App Protection";        U="https://graph.microsoft.com/beta/deviceAppManagement/managedAppPolicies?`$select=id,displayName" }
    )
    # improved reporting with included missing intune parts from the original 1.0 reporting script.
    $eps += @(
        # Windows Update
        @{ L="Quality Update"; U="https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Windows Update Ring"; U="https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$filter=isof('microsoft.graph.windowsUpdateForBusinessConfiguration')&`$expand=assignments&`$select=id,displayName,assignments" }

        # Autopilot and enrollment
        @{ L="Autopilot Profile"; U="https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Autopilot Device"; U="https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$select=id,serialNumber,productKey,azureAdDeviceId,groupTag,profileAssignedDateTime,deploymentProfileAssignmentStatus" }
        @{ L="Terms and Conditions"; U="https://graph.microsoft.com/beta/deviceManagement/termsAndConditions?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Device Category"; U="https://graph.microsoft.com/beta/deviceManagement/deviceCategories?`$select=id,displayName,description" }

        # Endpoint security, baselines, and intents
        @{ L="Endpoint Security Intent"; U="https://graph.microsoft.com/beta/deviceManagement/intents?`$select=id,displayName,description,isAssigned,templateId" }
        @{ L="Security Baseline"; U="https://graph.microsoft.com/beta/deviceManagement/intents?`$filter=templateId ne null&`$select=id,displayName,description,isAssigned,templateId" }

        # Apps
        @{ L="iOS App Provisioning"; U="https://graph.microsoft.com/beta/deviceAppManagement/iosLobAppProvisioningConfigurations?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="macOS App Provisioning"; U="https://graph.microsoft.com/beta/deviceAppManagement/macOsLobAppProvisioningConfigurations?`$select=id,displayName" }
        @{ L="Microsoft Store App"; U="https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.winGetApp')&`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Web App"; U="https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.webApp')&`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="Android App"; U="https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.androidLobApp') or isof('microsoft.graph.managedAndroidStoreApp') or isof('microsoft.graph.androidStoreApp')&`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="iOS App"; U="https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.iosLobApp') or isof('microsoft.graph.iosStoreApp') or isof('microsoft.graph.iosVppApp')&`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="macOS App"; U="https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.macOSLobApp') or isof('microsoft.graph.macOSDmgApp') or isof('microsoft.graph.macOSPkgApp')&`$expand=assignments&`$select=id,displayName,assignments" }

        # Scripts
        @{ L="macOS Shell Script"; U="https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts?`$expand=assignments&`$select=id,displayName,assignments" }
        @{ L="macOS Custom Attribute"; U="https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts?`$expand=assignments&`$select=id,displayName,assignments" }

        # Governance and targeting
        @{ L="Assignment Filter"; U="https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?`$select=id,displayName,description,platform,rule" }
        @{ L="Policy Set"; U="https://graph.microsoft.com/beta/deviceAppManagement/policySets?`$select=id,displayName,description,assignmentStatus" }
        @{ L="Intune Role Definition"; U="https://graph.microsoft.com/beta/deviceManagement/roleDefinitions?`$select=id,displayName,description,isBuiltIn" }
        @{ L="Intune Role Assignment"; U="https://graph.microsoft.com/beta/deviceManagement/roleAssignments?`$expand=microsoft.graph.deviceAndAppManagementRoleAssignment/assignments(`$expand=microsoft.graph.deviceAndAppManagementRoleAssignment/assignments)&`$select=id,displayName,description" }
        @{ L="Scope Tag"; U="https://graph.microsoft.com/beta/deviceManagement/scopedTags?`$select=id,displayName,description" }
    )

    foreach ($ep in $eps) {
        #ping localhost -n 30 # required to avoid throttling from MS servers
        try {
            $items = @(Invoke-GraphGetAll -Uri $ep.U -Silent)
            foreach ($item in $items) {
                $name  = (Get-AnyProp -Obj $item -Names 'displayName','name') ?? "Unnamed"
                $asgns = @(Get-AnyProp -Obj $item -Names 'assignments')
                foreach ($a in $asgns) {
                    if ($null -eq $a) { continue }
                    $target = Get-AnyProp -Obj $a -Names 'target'
                    $gid    = Get-AnyProp -Obj $target -Names 'groupId'
                    if ($gid) {
                        if (-not $assignIdx.ContainsKey($gid)) {
                            $assignIdx[$gid] = [System.Collections.Generic.List[psobject]]::new()
                        }
                        $assignIdx[$gid].Add("$($ep.L): $name")
                    }
                }
            }
            Write-Log "  Indexed $($ep.L): $($items.Count) items" INFO
        } catch {
            Write-Log "  Skipped '$($ep.L)': $($_.Exception.Message)" WARN
        }
    }

    Write-Log "Intune index done — $($assignIdx.Keys.Count) groups have assignments." SUCCESS
    return $assignIdx
}



#endregion

#region ══════════════════════════════════════════════════════
#  MAIN
#══════════════════════════════════════════════════════════════

Write-Log "====================================================" SECTION
Write-Log "  GROUP NESTING ANALYSIS — READ-ONLY v2.0 (PS7)"    SECTION
Write-Log "====================================================" SECTION
Write-Log "Output    : $OutputPath"    INFO
Write-Log "Threshold : depth >= $DeepNestingThreshold = HIGH risk" INFO

Write-Log "" INFO
Write-Log "STEP 0 — Connecting to Microsoft Graph" SECTION
try {
    Connect-MgGraph -Scopes $RequiredScopes -NoWelcome
    $ctx = Get-MgContext
    Write-Log "Connected: $($ctx.Account) | Tenant: $($ctx.TenantId)" SUCCESS
} catch {
    Write-Log "Connection failed: $($_.Exception.Message)" ERROR
    exit 1
}

try {

# ── STEP 1 ───────────────────────────────────────────────────────────────────
Write-Log "" INFO
Write-Log "STEP 1 — Retrieving all Entra ID groups" SECTION
$allGroups = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,description,groupTypes,createdDateTime,securityEnabled,mailEnabled,membershipRule,renewedDateTime")
Write-Log "Retrieved $($allGroups.Count) groups." SUCCESS

#ping localhost -n 60 # required to avoid throttling from MS servers

$groupLookup = @{}
$allGroups | ForEach-Object { $groupLookup[$_.id] = $_.displayName }
$groupIds = @($allGroups | ForEach-Object { $_.id })

#ping localhost -n 60 # required to avoid throttling from MS servers

# ── STEP 2 ───────────────────────────────────────────────────────────────────
Write-Log "" INFO
Write-Log "STEP 2 — Fetching direct members (batch GET)" SECTION
$directMap = Invoke-GraphBatch -GroupIds $groupIds -MemberEndpoint "members"
Write-Log "Direct members fetched for $($directMap.Keys.Count) groups." SUCCESS

#ping localhost -n 60 # required to avoid throttling from MS servers

# ── STEP 3 ───────────────────────────────────────────────────────────────────
Write-Log "" INFO
Write-Log "STEP 3 — Fetching transitive members (batch GET)" SECTION
$transitiveMap = Invoke-GraphBatch -GroupIds $groupIds -MemberEndpoint "transitiveMembers"
Write-Log "Transitive members fetched for $($transitiveMap.Keys.Count) groups." SUCCESS

#ping localhost -n 60 # required to avoid throttling from MS servers

# ── STEP 4 ───────────────────────────────────────────────────────────────────
Write-Log "" INFO
Write-Log "STEP 4 — Building Intune assignment index" SECTION
$intuneIdx = Get-IntuneAssignmentIndex

#ping localhost -n 60 # required to avoid throttling from MS servers

# ── STEP 5 ───────────────────────────────────────────────────────────────────
Write-Log "" INFO
Write-Log "STEP 5 — Building full nesting tree" SECTION

$nestingRows       = [System.Collections.Generic.List[object]]::new()
$groupsWithNesting = 0
$groupIdx          = 0

foreach ($group in $allGroups) {
    $groupIdx++
    Write-Progress -Activity "Building Nesting Tree" `
        -Status "[$groupIdx / $($allGroups.Count)] $($group.displayName)" `
        -PercentComplete ([math]::Round(($groupIdx / $allGroups.Count) * 100))
    try {
    $directMbrs = if ($directMap.ContainsKey($group.id)) { @($directMap[$group.id]) } else { @() }
    $childGrps  = @($directMbrs | Where-Object { (Get-OdataType $_) -eq '#microsoft.graph.group' })

    if ($childGrps.Count -gt 0) {
        $groupsWithNesting++
        $visited = [System.Collections.Generic.HashSet[string]]::new()
        $visited.Add($group.id) | Out-Null

        

            $rows = @(Build-NestingTree `
                -GroupId       $group.id `
                -RootId        $group.id `
                -Level         0 `
                -NestThreshold $DeepNestingThreshold `
                -DirectMap     $directMap `
                -GroupLookup   $groupLookup `
                -Visited       $visited `
                -Path          ($groupLookup[$group.id] ?? $group.id) `
                -PathIds       $group.id)
            if ($rows.Count -gt 0) { $nestingRows.AddRange([object[]]$rows) }
        }             
    } catch {
            
            Write-Log "Nesting tree error for group '$($group.displayName)' ($($group.id)): $($_.Exception.Message)" WARN
            }
}
Write-Progress -Activity "Building Nesting Tree" -Completed
Write-Log "Nesting tree: $($nestingRows.Count) relationships across $groupsWithNesting root groups." SUCCESS

# ── STEP 6 ───────────────────────────────────────────────────────────────────
Write-Log "" INFO
Write-Log "STEP 6 — Pre-building nesting role lookups" SECTION

$parentGroupIds    = [System.Collections.Generic.HashSet[string]]::new()
$childGroupIds     = [System.Collections.Generic.HashSet[string]]::new()
$maxDepthByRoot    = @{}
$hasCircularByRoot = @{}

foreach ($nr in $nestingRows) {
    $parentGroupIds.Add($nr.RootGroupId)   | Out-Null
    $parentGroupIds.Add($nr.ParentGroupId) | Out-Null
    $childGroupIds.Add($nr.ChildGroupId)   | Out-Null

    if (-not $nr.IsCircular) {
        $cur = if ($maxDepthByRoot.ContainsKey($nr.RootGroupId)) { $maxDepthByRoot[$nr.RootGroupId] } else { 0 }
        if ($nr.NestingDepth -gt $cur) { $maxDepthByRoot[$nr.RootGroupId] = $nr.NestingDepth }
    }
    if ($nr.IsCircular) { $hasCircularByRoot[$nr.RootGroupId] = $true }
}
Write-Log "Lookup sets built." SUCCESS

# ── STEP 6b ──────────────────────────────────────────────────────────────────
Write-Log "" INFO
Write-Log "STEP 6b — Building per-group summary table" SECTION

$groupSummary = [System.Collections.Generic.List[object]]::new()

foreach ($group in $allGroups) {
    $directMbrs     = if ($directMap.ContainsKey($group.id))     { @($directMap[$group.id]) }     else { @() }
    $transitiveMbrs = if ($transitiveMap.ContainsKey($group.id)) { @($transitiveMap[$group.id]) } else { @() }

    $directUsers   = @($directMbrs     | Where-Object { (Get-OdataType $_) -eq '#microsoft.graph.user' })
    $directDevices = @($directMbrs     | Where-Object { (Get-OdataType $_) -eq '#microsoft.graph.device' })
    $directGroups  = @($directMbrs     | Where-Object { (Get-OdataType $_) -eq '#microsoft.graph.group' })
    $transUsers    = @($transitiveMbrs | Where-Object { (Get-OdataType $_) -eq '#microsoft.graph.user' })
    $transDevices  = @($transitiveMbrs | Where-Object { (Get-OdataType $_) -eq '#microsoft.graph.device' })

    $safeGroupTypes = @($group.groupTypes)
    $isDynamic = $safeGroupTypes -contains "DynamicMembership"
    $isM365    = $safeGroupTypes -contains "Unified"

    $groupType = if     ($isM365 -and $isDynamic)                             { "Dynamic M365" }
                elseif ($isM365)                                             { "Microsoft 365" }
                elseif ($isDynamic)                                          { "Dynamic Security" }
                elseif ($group.mailEnabled -and -not $group.securityEnabled) { "Distribution" }
                else                                                          { "Security" }

    $isParentRole = $parentGroupIds.Contains($group.id)
    $isChildRole  = $childGroupIds.Contains($group.id)
    $nestingRole  = if     ($isParentRole -and $isChildRole) { "Parent AND Child" }
                    elseif ($isParentRole)                   { "Parent (contains nested groups)" }
                    elseif ($isChildRole)                    { "Child (nested inside another group)" }
                    else                                     { "No nesting" }

    $maxDepth    = if ($maxDepthByRoot.ContainsKey($group.id))    { $maxDepthByRoot[$group.id] }    else { 0 }
    $hasCircular = if ($hasCircularByRoot.ContainsKey($group.id)) { $hasCircularByRoot[$group.id] } else { $false }

    $intuneAssignments = if ($intuneIdx.ContainsKey($group.id)) { ($intuneIdx[$group.id] -join " | ") } else { "" }
    $intuneCount       = if ($intuneIdx.ContainsKey($group.id)) { $intuneIdx[$group.id].Count }       else { 0 }

    $memberInflation = if ($directMbrs.Count -gt 0 -and $transitiveMbrs.Count -gt $directMbrs.Count) {
        "$($transitiveMbrs.Count - $directMbrs.Count) extra members via nesting"
    } elseif ($directMbrs.Count -eq 0 -and $transitiveMbrs.Count -gt 0) {
        "Dynamic — $($transitiveMbrs.Count) transitive members"
    } else { "" }

    $groupSummary.Add([PSCustomObject]@{
        GroupId                = $group.id
        GroupName              = $group.displayName
        GroupType              = $groupType
        IsDynamic              = $isDynamic
        Description            = $group.description
        CreatedDate            = $group.createdDateTime
        DirectMembersTotal     = $directMbrs.Count
        DirectUsers            = $directUsers.Count
        DirectDevices          = $directDevices.Count
        DirectNestedGroups     = $directGroups.Count
        TransitiveMembersTotal = $transitiveMbrs.Count
        TransitiveUsers        = $transUsers.Count
        TransitiveDevices      = $transDevices.Count
        MemberInflationNote    = $memberInflation
        NestingRole            = $nestingRole
        MaxNestingDepthAsRoot  = $maxDepth
        HasCircularReference   = $hasCircular
        IntuneAssignmentCount  = $intuneCount
        HasIntuneAssignment    = ($intuneCount -gt 0)
        IntuneAssignments      = $intuneAssignments
    })
}
Write-Log "Group summary: $($groupSummary.Count) rows." SUCCESS

# ── STEP 7 ───────────────────────────────────────────────────────────────────
Write-Log "" INFO
Write-Log "STEP 7 — Building recommendations" SECTION

$recoRows      = [System.Collections.Generic.List[object]]::new()
$recoProcessed = [System.Collections.Generic.HashSet[string]]::new()

# Circular references
$circularGroups = @($nestingRows | Where-Object { $_.IsCircular } |
    ForEach-Object { $_.RootGroupId } | Select-Object -Unique)

foreach ($rootId in $circularGroups) {
    $circRows    = @($nestingRows | Where-Object { $_.IsCircular -and $_.RootGroupId -eq $rootId })
    $rootName    = $groupLookup[$rootId] ?? $rootId
    $intuneCtx   = if ($intuneIdx.ContainsKey($rootId)) { ($intuneIdx[$rootId] -join "; ") } else { "No Intune assignment" }
    $affectedIds = ($circRows | ForEach-Object { $_.ChildGroupId }) -join ", "

    $recoRows.Add([PSCustomObject]@{
        Priority            = "1 - CRITICAL"
        RootGroupId         = $rootId
        RootGroupName       = $rootName
        IssueType           = "CIRCULAR REFERENCE"
        NestingDepth        = "N/A"
        AffectedChildGroups = $affectedIds
        IntuneContext       = $intuneCtx
        FullPathChain       = ($circRows | ForEach-Object { $_.PathChain }) -join " || "
        Recommendation      = "URGENT - FIX FIRST: Circular nesting causes infinite Intune policy loops. (1) Open Entra > Groups > '$rootName' > Members. (2) Find and remove the member group creating the loop. (3) Re-run script to confirm. Affected IDs: $affectedIds"
        EffortEstimate      = "15-30 min"
        BusinessRisk        = "CRITICAL - Auth failures and incorrect policy application possible"
    })
    $recoProcessed.Add($rootId) | Out-Null
}

# Deep nesting
$deepRows  = @($nestingRows | Where-Object { -not $_.IsCircular -and $_.NestingDepth -ge $DeepNestingThreshold })
$deepRoots = @($deepRows | Select-Object -ExpandProperty RootGroupId -Unique)

foreach ($rootId in $deepRoots) {
    if ($recoProcessed.Contains($rootId)) { continue }
    $myDeep    = @($deepRows | Where-Object { $_.RootGroupId -eq $rootId })
    $maxD      = ($myDeep | Measure-Object -Property NestingDepth -Maximum).Maximum
    $rootName  = $groupLookup[$rootId] ?? $rootId
    $intuneCtx = if ($intuneIdx.ContainsKey($rootId)) { ($intuneIdx[$rootId] -join "; ") } else { "No Intune assignment" }
    $worstPath = ($myDeep | Sort-Object NestingDepth -Descending | Select-Object -First 1).PathChain

    $recoRows.Add([PSCustomObject]@{
        Priority            = "2 - HIGH"
        RootGroupId         = $rootId
        RootGroupName       = $rootName
        IssueType           = "DEEP NESTING (depth $maxD)"
        NestingDepth        = $maxD
        AffectedChildGroups = ($myDeep | Select-Object -ExpandProperty ChildGroupId -Unique) -join ", "
        IntuneContext       = $intuneCtx
        FullPathChain       = $worstPath
        Recommendation      = "FLATTEN: $maxD levels creates overhead on every Intune policy check. Review: $worstPath. Add members of deepest groups directly to '$rootName' or the policy target, then remove intermediate layers."
        EffortEstimate      = "30-60 min per root group"
        BusinessRisk        = "HIGH - Policy evaluation overhead, hard-to-troubleshoot scope"
    })
    $recoProcessed.Add($rootId) | Out-Null
}

# Dynamic group nested as child
$dynamicChildGroups = @($groupSummary | Where-Object {
    $_.IsDynamic -eq $true -and
    $_.NestingRole -in @("Child (nested inside another group)", "Parent AND Child")
})
foreach ($g in $dynamicChildGroups) {
    if ($recoProcessed.Contains($g.GroupId)) { continue }
    $intuneCtx = if ($intuneIdx.ContainsKey($g.GroupId)) { ($intuneIdx[$g.GroupId] -join "; ") } else { "No Intune assignment" }
    $recoRows.Add([PSCustomObject]@{
        Priority            = "2b - HIGH"
        RootGroupId         = $g.GroupId
        RootGroupName       = $g.GroupName
        IssueType           = "DYNAMIC GROUP NESTED AS CHILD"
        NestingDepth        = $g.MaxNestingDepthAsRoot
        AffectedChildGroups = ""
        IntuneContext       = $intuneCtx
        FullPathChain       = ""
        Recommendation      = "REVIEW: '$($g.GroupName)' is a dynamic group nested inside another group. Entra evaluates dynamic rules independently — the parent group's Intune policy scope may NOT reflect the dynamic rule. Assign Intune policies directly to this group instead."
        EffortEstimate      = "15-30 min"
        BusinessRisk        = "HIGH - Unpredictable Intune policy scope"
    })
}

# Double nesting
$mediumRows = @($nestingRows | Where-Object {
    -not $_.IsCircular -and
    $_.NestingDepth -eq 2 -and
    -not $recoProcessed.Contains($_.RootGroupId)
})
$mediumRoots = @($mediumRows | Select-Object -ExpandProperty RootGroupId -Unique)

foreach ($rootId in $mediumRoots) {
    if ($recoProcessed.Contains($rootId)) { continue }
    $rootName    = $groupLookup[$rootId] ?? $rootId
    $intuneCtx   = if ($intuneIdx.ContainsKey($rootId)) { ($intuneIdx[$rootId] -join "; ") } else { "No Intune assignment" }
    $myRows      = @($mediumRows | Where-Object { $_.RootGroupId -eq $rootId })
    $firstParent = if ($myRows.Count -gt 0) { $myRows[0].ParentGroupName } else { "unknown" }

    $recoRows.Add([PSCustomObject]@{
        Priority            = "3 - MEDIUM"
        RootGroupId         = $rootId
        RootGroupName       = $rootName
        IssueType           = "DOUBLE NESTING (depth 2)"
        NestingDepth        = 2
        AffectedChildGroups = ($myRows | Select-Object -ExpandProperty ChildGroupId -Unique) -join ", "
        IntuneContext       = $intuneCtx
        FullPathChain       = if ($myRows.Count -gt 0) { $myRows[0].PathChain } else { "" }
        Recommendation      = "REVIEW: Check if intermediate group '$firstParent' serves any purpose beyond being a container. If not, move its members directly into '$rootName' and remove the layer. If it does, document it in the group description."
        EffortEstimate      = "15-30 min"
        BusinessRisk        = "LOW - Minor complexity overhead"
    })
    $recoProcessed.Add($rootId) | Out-Null
}

# Single nesting, no Intune assignment
$lowRows = @($nestingRows | Where-Object {
    -not $_.IsCircular -and
    $_.NestingDepth -eq 1 -and
    -not $recoProcessed.Contains($_.RootGroupId) -and
    -not $intuneIdx.ContainsKey($_.RootGroupId)
})
$lowRoots = @($lowRows | Select-Object -ExpandProperty RootGroupId -Unique)

foreach ($rootId in $lowRoots) {
    if ($recoProcessed.Contains($rootId)) { continue }
    $rootName = $groupLookup[$rootId] ?? $rootId
    $myRows   = @($lowRows | Where-Object { $_.RootGroupId -eq $rootId })

    $recoRows.Add([PSCustomObject]@{
        Priority            = "4 - LOW"
        RootGroupId         = $rootId
        RootGroupName       = $rootName
        IssueType           = "SINGLE NESTING - No Intune assignment"
        NestingDepth        = 1
        AffectedChildGroups = ""
        IntuneContext       = "No Intune assignment found"
        FullPathChain       = if ($myRows.Count -gt 0) { $myRows[0].PathChain } else { "" }
        Recommendation      = "MONITOR: Single-level nesting with no Intune assignment. Likely used for access control (SharePoint/Teams). Review nesting before assigning to any Intune policy in future."
        EffortEstimate      = "5 min review"
        BusinessRisk        = "INFORMATIONAL"
    })
    $recoProcessed.Add($rootId) | Out-Null
}

# Empty groups with Intune assignments
$emptyAssigned = @($groupSummary | Where-Object {
    $_.DirectMembersTotal     -eq 0 -and
    $_.TransitiveMembersTotal -eq 0 -and
    $_.HasIntuneAssignment    -eq $true
})
foreach ($g in $emptyAssigned) {
    $recoRows.Add([PSCustomObject]@{
        Priority            = "5 - MEDIUM"
        RootGroupId         = $g.GroupId
        RootGroupName       = $g.GroupName
        IssueType           = "EMPTY GROUP WITH INTUNE ASSIGNMENT"
        NestingDepth        = 0
        AffectedChildGroups = ""
        IntuneContext       = $g.IntuneAssignments
        FullPathChain       = ""
        Recommendation      = "CLEAN UP: '$($g.GroupName)' has zero members but is assigned to Intune policies. Either add the intended members, or remove the assignments and archive/delete this group."
        EffortEstimate      = "10 min"
        BusinessRisk        = "MEDIUM - Wasted policy evaluation, misleading scope"
    })
}

Write-Log "Recommendations: $($recoRows.Count) rows." SUCCESS

# ── STEP 8 ───────────────────────────────────────────────────────────────────
$totalNestingRels  = $nestingRows.Count
$circularCount     = (@($nestingRows | Where-Object { $_.IsCircular })).Count
$deepCount         = (@($nestingRows | Where-Object { -not $_.IsCircular -and $_.NestingDepth -ge $DeepNestingThreshold })).Count
$nonCircRows       = @($nestingRows | Where-Object { -not $_.IsCircular })
$maxDepthOverall   = if ($nonCircRows.Count -gt 0) { ($nonCircRows | Measure-Object -Property NestingDepth -Maximum).Maximum } else { 0 }
$groupsAsParent    = $parentGroupIds.Count
$groupsAsChildOnly = (@($groupSummary | Where-Object { $_.NestingRole -eq "Child (nested inside another group)" })).Count
$emptyWithAssign   = $emptyAssigned.Count

#endregion

#region ══════════════════════════════════════════════════════
#  EXCEL EXPORT
#══════════════════════════════════════════════════════════════

Write-Log "" INFO
Write-Log "STEP 9 — Exporting Excel workbook" SECTION

$dir = Split-Path $OutputPath -Parent
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$cfCritical = New-ConditionalText -Text "CRITICAL" -BackgroundColor "DarkRed"     -ConditionalTextColor "White"
$cfHigh     = New-ConditionalText -Text "HIGH"     -BackgroundColor "OrangeRed"   -ConditionalTextColor "White"
$cfMedium   = New-ConditionalText -Text "MEDIUM"   -BackgroundColor "Orange"      -ConditionalTextColor "Black"
$cfLow      = New-ConditionalText -Text "LOW"      -BackgroundColor "Gold"        -ConditionalTextColor "Black"
$cfCircTrue = New-ConditionalText -Text "True"     -BackgroundColor "Red"         -ConditionalTextColor "White"
$cfParent   = New-ConditionalText -Text "Parent"   -BackgroundColor "LightBlue"   -ConditionalTextColor "Black"
$cfChild    = New-ConditionalText -Text "Child"    -BackgroundColor "LightYellow" -ConditionalTextColor "Black"
$cfInfo     = New-ConditionalText -Text "INFORM"   -BackgroundColor "LightGreen"  -ConditionalTextColor "Black"
$cfDynamic  = New-ConditionalText -Text "Dynamic"  -BackgroundColor "Lavender"    -ConditionalTextColor "Black"

# SHEET 1 — DASHBOARD
Write-Log "  Sheet 1 — Dashboard..." INFO
$dashRows = @(
    [PSCustomObject]@{ Category="REPORT INFO";     Metric="Generated";                       Value=(Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
    [PSCustomObject]@{ Category="";               Metric="Tenant ID";                       Value=$ctx.TenantId }
    [PSCustomObject]@{ Category="";               Metric="Admin Account";                   Value=$ctx.Account }
    [PSCustomObject]@{ Category="";               Metric="PowerShell Version";              Value=$PSVersionTable.PSVersion.ToString() }
    [PSCustomObject]@{ Category="";               Metric="Deep Nesting Threshold";          Value="Depth >= $DeepNestingThreshold = HIGH risk" }
    [PSCustomObject]@{ Category="------";         Metric="";                               Value="" }
    [PSCustomObject]@{ Category="INVENTORY";       Metric="Total Groups";                   Value=$allGroups.Count }
    [PSCustomObject]@{ Category="";               Metric="Groups containing nested groups"; Value=$groupsAsParent }
    [PSCustomObject]@{ Category="";               Metric="Groups nested inside others";     Value=$groupsAsChildOnly }
    [PSCustomObject]@{ Category="";               Metric="Total Nesting Relationships";     Value=$totalNestingRels }
    [PSCustomObject]@{ Category="";               Metric="Maximum Nesting Depth Found";     Value=$maxDepthOverall }
    [PSCustomObject]@{ Category="";               Metric="Empty Groups with Assignments";   Value=$emptyWithAssign }
    [PSCustomObject]@{ Category="------";         Metric="";                               Value="" }
    [PSCustomObject]@{ Category="CRITICAL ISSUES"; Metric="Circular References";            Value=$circularCount }
    [PSCustomObject]@{ Category="";               Metric="Deep Nesting (>= depth $DeepNestingThreshold)"; Value=$deepCount }
    [PSCustomObject]@{ Category="------";         Metric="";                               Value="" }
    [PSCustomObject]@{ Category="ACTION SUMMARY";  Metric="Priority 1 CRITICAL";            Value=(@($recoRows | Where-Object { $_.Priority -like "1*" })).Count }
    [PSCustomObject]@{ Category="";               Metric="Priority 2 HIGH (nesting)";      Value=(@($recoRows | Where-Object { $_.Priority -like "2 -*" })).Count }
    [PSCustomObject]@{ Category="";               Metric="Priority 2b HIGH (dynamic)";     Value=(@($recoRows | Where-Object { $_.Priority -like "2b*" })).Count }
    [PSCustomObject]@{ Category="";               Metric="Priority 3 MEDIUM";              Value=(@($recoRows | Where-Object { $_.Priority -like "3*" })).Count }
    [PSCustomObject]@{ Category="";               Metric="Priority 4 LOW";                 Value=(@($recoRows | Where-Object { $_.Priority -like "4*" })).Count }
    [PSCustomObject]@{ Category="";               Metric="Priority 5 Empty Groups";        Value=(@($recoRows | Where-Object { $_.Priority -like "5*" })).Count }
    [PSCustomObject]@{ Category="------";         Metric="";                               Value="" }
    [PSCustomObject]@{ Category="HOW TO READ";     Metric="Sheet 2 - Nesting Tree";         Value="Full hierarchy: every parent->child link, depth, risk colour" }
    [PSCustomObject]@{ Category="";               Metric="Sheet 3 - All Groups";            Value="Every group: member counts, type, nesting role, Intune assignments" }
    [PSCustomObject]@{ Category="";               Metric="Sheet 4 - Circular References";   Value="Fix FIRST — causes policy evaluation loops" }
    [PSCustomObject]@{ Category="";               Metric="Sheet 5 - Deep Nesting";          Value="All relationships >= depth $DeepNestingThreshold, sorted deepest first" }
    [PSCustomObject]@{ Category="";               Metric="Sheet 6 - Recommendations";       Value="All findings sorted by priority — use as action plan" }
    [PSCustomObject]@{ Category="";               Metric="Sheet 7 - Risk Pivot";            Value="Risk distribution chart" }
)
$dashRows | Export-Excel -Path $OutputPath -WorksheetName "Dashboard" `
    -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Light15

# SHEET 2 — NESTING TREE
Write-Log "  Sheet 2 — Nesting Tree ($($nestingRows.Count) rows)..." INFO
if ($nestingRows.Count -gt 0) {
    $nestingRows | Select-Object `
        RootGroupName, RootGroupId, ParentGroupName, ParentGroupId,
        ChildGroupName, ChildGroupId, NestingDepth, PathChain,
        IsCircular, RiskLevel, RiskReason |
    Export-Excel -Path $OutputPath -WorksheetName "Nesting Tree" `
        -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
        -ConditionalFormat $cfCritical,$cfHigh,$cfMedium,$cfLow,$cfCircTrue `
        -TableStyle Medium6
} else {
    @([PSCustomObject]@{ Note = "No group nesting found in this tenant." }) |
    Export-Excel -Path $OutputPath -WorksheetName "Nesting Tree" -AutoSize
}

# SHEET 3 — ALL GROUPS
Write-Log "  Sheet 3 — All Groups ($($groupSummary.Count) rows)..." INFO
$groupSummary | Select-Object `
    GroupName, GroupId, GroupType, IsDynamic, Description, CreatedDate,
    DirectMembersTotal, DirectUsers, DirectDevices, DirectNestedGroups,
    TransitiveMembersTotal, TransitiveUsers, TransitiveDevices,
    MemberInflationNote, NestingRole, MaxNestingDepthAsRoot, HasCircularReference,
    IntuneAssignmentCount, HasIntuneAssignment, IntuneAssignments |
Export-Excel -Path $OutputPath -WorksheetName "All Groups" `
    -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
    -ConditionalFormat $cfCircTrue,$cfParent,$cfChild,$cfDynamic `
    -TableStyle Medium2

# SHEET 4 — CIRCULAR REFERENCES
Write-Log "  Sheet 4 — Circular References..." INFO
$circularOnly = @($nestingRows | Where-Object { $_.IsCircular })
if ($circularOnly.Count -gt 0) {
    $circularOnly | Select-Object `
        RootGroupName, RootGroupId, ParentGroupName, ParentGroupId,
        ChildGroupName, ChildGroupId, NestingDepth, PathChain,
        RiskLevel, RiskReason, Recommendation |
    Export-Excel -Path $OutputPath -WorksheetName "Circular References" `
        -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
        -ConditionalFormat $cfCritical -TableStyle Medium14
} else {
    @([PSCustomObject]@{ Result = "No circular references found. Nesting structure is loop-free." }) |
    Export-Excel -Path $OutputPath -WorksheetName "Circular References" -AutoSize
}

# SHEET 5 — DEEP NESTING
Write-Log "  Sheet 5 — Deep Nesting..." INFO
$deepOnly = @($nestingRows | Where-Object { -not $_.IsCircular -and $_.NestingDepth -ge $DeepNestingThreshold })
if ($deepOnly.Count -gt 0) {
    $deepOnly | Sort-Object NestingDepth -Descending |
    Select-Object `
        RootGroupName, RootGroupId, ParentGroupName, ParentGroupId,
        ChildGroupName, ChildGroupId, NestingDepth, PathChain,
        RiskLevel, RiskReason, Recommendation |
    Export-Excel -Path $OutputPath -WorksheetName "Deep Nesting" `
        -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
        -ConditionalFormat $cfCritical,$cfHigh -TableStyle Medium10
} else {
    @([PSCustomObject]@{ Result = "No nesting deeper than $($DeepNestingThreshold - 1) levels found." }) |
    Export-Excel -Path $OutputPath -WorksheetName "Deep Nesting" -AutoSize
}

# SHEET 6 — RECOMMENDATIONS
Write-Log "  Sheet 6 — Recommendations ($($recoRows.Count) rows)..." INFO
if ($recoRows.Count -gt 0) {
    $recoRows | Sort-Object Priority |
    Select-Object `
        Priority, RootGroupName, RootGroupId, IssueType, NestingDepth,
        IntuneContext, FullPathChain, Recommendation, EffortEstimate, BusinessRisk |
    Export-Excel -Path $OutputPath -WorksheetName "Recommendations" `
        -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
        -ConditionalFormat $cfCritical,$cfHigh,$cfMedium,$cfLow,$cfInfo `
        -TableStyle Light9
} else {
    @([PSCustomObject]@{ Result = "No issues found requiring action." }) |
    Export-Excel -Path $OutputPath -WorksheetName "Recommendations" -AutoSize
}



Write-Log "Excel export complete." SUCCESS

$xlsxNesting = $OutputPath -replace '\.xlsx$', '_NestingTree.xlsx'
$xlsxRecos   = $OutputPath -replace '\.xlsx$', '_Recommendations.xlsx'
$nestingRows | Export-Excel -Path $xlsxNesting -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
$recoRows    | Export-Excel -Path $xlsxRecos -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow

Write-Log "" INFO
Write-Log "====================================================" SECTION
Write-Log "  COMPLETE" SECTION
Write-Log "====================================================" SECTION
Write-Log "Excel : $OutputPath" SUCCESS
Write-Log "CSV 1 : $xlsxNesting" SUCCESS
Write-Log "CSV 2 : $xlsxRecos"   SUCCESS
Write-Log "" INFO
Write-Log "START HERE:" INFO
Write-Log "  1. Sheet 'Circular References' — fix FIRST (policy loops)" WARN
Write-Log "  2. Sheet 'Recommendations'     — sorted by priority"       INFO
Write-Log "  3. Sheet 'Nesting Tree'        — full visual hierarchy"     INFO

#endregion

} finally {
    try { Disconnect-MgGraph | Out-Null } catch { }
}
