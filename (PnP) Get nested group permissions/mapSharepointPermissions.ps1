Import-Module PnP.PowerShell -ErrorAction Stop

$ThumbPrint = "####"
$ClientId   = "####"
$TenantId   = "####"
$AdminUrl   = "https://<domain>-admin.sharepoint.com"

# ── Connect to admin to get all sites ────────────────────────────────────────
Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId
Write-Host "Connected to SharePoint Admin" -ForegroundColor Green

$allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false | Select-Object -ExpandProperty Url
Write-Host "Found $($allSites.Count) sites" -ForegroundColor Cyan

# ── Graph setup ───────────────────────────────────────────────────────────────
$graphToken   = (Get-PnPAccessToken -ResourceTypeName Graph)
$graphHeaders = @{ Authorization = "Bearer $graphToken" }

function Invoke-Graph {
    param([string]$Path)
    $url = "https://graph.microsoft.com/v1.0$Path"
    $all = @()
    do {
        $resp = Invoke-RestMethod $url -Headers $graphHeaders
        $all += $resp.value
        $url  = $resp.'@odata.nextLink'
    } while ($url)
    return $all
}

function Get-RuleSourceGroup {
    param([string]$MembershipRule)
    if ($MembershipRule -match '"([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})"') {
        $ruleGroupId = $Matches[1]
        try {
            $ruleGroup = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$ruleGroupId" -Headers $graphHeaders
            return [PSCustomObject]@{ Id = $ruleGroupId; DisplayName = $ruleGroup.displayName }
        } catch {
            return [PSCustomObject]@{ Id = $ruleGroupId; DisplayName = "(could not resolve)" }
        }
    }
    return $null
}

function Get-RuleJobTitles {
    param([string]$MembershipRule)
    $titles = [System.Collections.Generic.List[string]]::new()
    $found  = [regex]::Matches($MembershipRule, 'user\.jobTitle -eq "([^"]+)"')
    foreach ($m in $found) {
        $titles.Add($m.Groups[1].Value)
    }
    return ($titles -join ', ')
}

# ── Caches ────────────────────────────────────────────────────────────────────
$aadGroupCache       = @{}
$sourceGroupRuleCache = @{}

function Get-AadGroup {
    param([string]$Id)
    if (-not $aadGroupCache.ContainsKey($Id)) {
        $aadGroupCache[$Id] = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$Id" -Headers $graphHeaders
    }
    return $aadGroupCache[$Id]
}

# ── Process each site ─────────────────────────────────────────────────────────
$allResults = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($siteUrl in $allSites) {
    Write-Host "`nProcessing: $siteUrl" -ForegroundColor Cyan

    try {
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Thumbprint $ThumbPrint -Tenant $TenantId -ErrorAction Stop
    } catch {
        Write-Host "  [SKIP] Could not connect: $($_.Exception.Message)" -ForegroundColor Red
        continue
    }

    $ctx = Get-PnPContext
    $web = $ctx.Web
    $ctx.Load($web)
    $ctx.Load($web.RoleAssignments)
    $ctx.ExecuteQuery()

    foreach ($role in $web.RoleAssignments) {
        $ctx.Load($role.Member)
        $ctx.Load($role.RoleDefinitionBindings)
        $ctx.ExecuteQuery()

        $spGroupName = $role.Member.Title
        $permission  = ($role.RoleDefinitionBindings |
                        Where-Object { $_.Name -ne 'Limited Access' } |
                        Select-Object -ExpandProperty Name) -join ', '

        try { $members = Get-PnPGroupMember -Group $spGroupName -ErrorAction Stop }
        catch { continue }

        foreach ($member in $members) {
            $principalType = [string]$member.PrincipalType
            if ($principalType -notin @('SecurityGroup', 'Microsoft365Group')) { continue }

            $aadId = ($member.LoginName -replace '.*\|') -replace '_o$'
            if ($aadId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') { continue }

            try { $aadGroup = Get-AadGroup -Id $aadId } catch { continue }

            $ruleSourceGroup = Get-RuleSourceGroup -MembershipRule $aadGroup.membershipRule

            # Resolve source group's membership rule (with cache)
            $sourceGroupRule = ""
            if ($ruleSourceGroup) {
                if (-not $sourceGroupRuleCache.ContainsKey($ruleSourceGroup.Id)) {
                    try {
                        $sg = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$($ruleSourceGroup.Id)" -Headers $graphHeaders
                        $sourceGroupRuleCache[$ruleSourceGroup.Id] = $sg.membershipRule
                    } catch {
                        $sourceGroupRuleCache[$ruleSourceGroup.Id] = ""
                    }
                }
                $sourceGroupRule = $sourceGroupRuleCache[$ruleSourceGroup.Id]
            }

            $jobTitles = Get-RuleJobTitles -MembershipRule $sourceGroupRule

            try { $aadMembers = Invoke-Graph "/groups/$aadId/transitiveMembers" } catch { continue }

            foreach ($aadMember in $aadMembers) {
                $type = $aadMember.'@odata.type' -replace '#microsoft.graph.', ''
                if ($type -ne 'user') { continue }

                $allResults.Add([PSCustomObject]@{
                    SiteUrl       = $siteUrl
                    SPGroup       = $spGroupName
                    Permission    = $permission
                    M365Group     = $aadGroup.displayName
                    M365GroupId   = $aadId
                    SourceGroup   = if ($ruleSourceGroup) { $ruleSourceGroup.DisplayName } else { "(none)" }
                    SourceGroupId = if ($ruleSourceGroup) { $ruleSourceGroup.Id } else { "(none)" }
                    JobTitles     = $jobTitles
                    MemberName    = $aadMember.displayName
                    MemberUPN     = $aadMember.userPrincipalName
                })
            }
        }
    }
}

# ── Deduplicate per site + user ───────────────────────────────────────────────
$deduplicated = $allResults |
    Group-Object SiteUrl, MemberUPN |
    ForEach-Object {
        $first = $_.Group[0]
        [PSCustomObject]@{
            SiteUrl       = $first.SiteUrl
            MemberName    = $first.MemberName
            MemberUPN     = $first.MemberUPN
            Permissions   = ($_.Group.Permission | Sort-Object -Unique) -join ', '
            M365Group     = $first.M365Group
            M365GroupId   = $first.M365GroupId
            SourceGroup   = $first.SourceGroup
            SourceGroupId = $first.SourceGroupId
            JobTitles     = $first.JobTitles
        }
    }

Write-Host "`n=== FULL TENANT MEMBERSHIP CHAIN ===" -ForegroundColor Green
$deduplicated | Sort-Object SiteUrl, MemberName | Format-Table -AutoSize -Wrap

$deduplicated | Sort-Object SiteUrl, MemberName | Export-Csv -Path ".\tenant_permissions.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Exported to tenant_permissions.csv" -ForegroundColor Green