function Set-CIPPDBCacheSharePointSharingLinks {
    <#
    .SYNOPSIS
        Caches SharePoint and OneDrive sharing links for a tenant

    .DESCRIPTION
        Enumerates every drive in the tenant (SharePoint sites and OneDrive personal sites),
        scans each drive via the driveItem delta endpoint for items carrying the "shared" facet
        and collects the direct (non-inherited) sharing permissions of those items. Sharing
        links are classified as Anonymous, External or Internal based on the link scope and the
        recipients' domains; direct (non-link) grants to external users are included as well.

    .PARAMETER TenantFilter
        The tenant to cache sharing links for

    .PARAMETER QueueId
        The queue ID to update with total tasks (optional)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantFilter,
        [string]$QueueId
    )

    # Returns $true when an identity (link recipient or direct grant) is external to the tenant.
    function Test-CIPPExternalIdentity {
        param($Identity, $InternalDomains)
        $LoginName = [string]($Identity.siteUser.loginName ?? $Identity.user.loginName ?? '')
        if ($LoginName -match '#ext#' -or $LoginName -match 'urn%3aspo%3aguest' -or $LoginName -match 'urn:spo:guest') { return $true }
        $Email = [string]($Identity.user.email ?? $Identity.user.userPrincipalName ?? $Identity.siteUser.email ?? '')
        if ($Email -match '#EXT#') { return $true }
        if ($Email -and $Email.Contains('@') -and $InternalDomains.Count -gt 0) {
            return -not $InternalDomains.Contains($Email.Split('@')[-1])
        }
        return $false
    }

    # Friendly display value for an identity, preferring email over display name.
    function Get-CIPPIdentityLabel {
        param($Identity)
        $Identity.user.email ?? $Identity.user.userPrincipalName ?? $Identity.siteUser.email ?? $Identity.user.displayName ?? $Identity.siteUser.displayName ?? $Identity.group.email ?? $Identity.group.displayName ?? $Identity.siteGroup.displayName
    }

    $StartTime = Get-Date
    Write-Host "[SharingLinks][$TenantFilter] START at $($StartTime.ToString('o'))"
    try {
        Write-LogMessage -API 'CIPPDBCache' -tenant $TenantFilter -message 'Caching SharePoint and OneDrive sharing links' -sev Debug

        # Verified domains, used to tell internal from external recipients.
        $InternalDomains = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        try {
            $Domains = New-GraphGetRequest -uri "https://graph.microsoft.com/v1.0/domains?`$select=id,isVerified" -tenantid $TenantFilter -asapp $true
            foreach ($Domain in ($Domains | Where-Object { $_.isVerified })) { [void]$InternalDomains.Add($Domain.id) }
            Write-Host "[SharingLinks][$TenantFilter] Phase 0 (domains): $($InternalDomains.Count) verified domains"
        } catch {
            Write-Host "[SharingLinks][$TenantFilter] Phase 0 (domains) FAILED: $($_.Exception.Message)"
            Write-LogMessage -API 'CIPPDBCache' -tenant $TenantFilter -message "Could not list verified domains for sharing link classification: $($_.Exception.Message)" -sev Warning
        }

        # 1) All sites, including OneDrive personal sites.
        try {
            $Sites = @(New-GraphGetRequest -uri "https://graph.microsoft.com/beta/sites/getAllSites?`$select=id,displayName,name,webUrl,isPersonalSite&`$top=999" -tenantid $TenantFilter -asapp $true)
        } catch {
            Write-Host "[SharingLinks][$TenantFilter] Phase 1 (getAllSites) FAILED: $($_.Exception.Message)"
            throw
        }
        Write-Host "[SharingLinks][$TenantFilter] Phase 1 (getAllSites): $($Sites.Count) sites (+$([math]::Round(((Get-Date) - $StartTime).TotalSeconds,1))s)"

        # 2) All document libraries (drives) per site, in bulk.
        $SiteByRequestId = @{}
        $RequestId = 0
        $DriveRequests = foreach ($Site in $Sites) {
            $SiteByRequestId["$RequestId"] = $Site
            @{
                id     = "$RequestId"
                method = 'GET'
                url    = "sites/$($Site.id)/drives?`$select=id,name,driveType,webUrl"
            }
            $RequestId++
        }

        $DriveEntries = [System.Collections.Generic.List[object]]::new()
        $FailedDriveLookups = 0
        if (@($DriveRequests).Count -gt 0) {
            Write-Host "[SharingLinks][$TenantFilter] Phase 2 (site drives): requesting drives for $(@($DriveRequests).Count) sites"
            try {
                $DriveResponses = New-GraphBulkRequest -tenantid $TenantFilter -Requests @($DriveRequests) -asapp $true
            } catch {
                Write-Host "[SharingLinks][$TenantFilter] Phase 2 (site drives) BULK FAILED: $($_.Exception.Message)"
                throw
            }
            foreach ($Response in $DriveResponses) {
                if ($Response.status -and $Response.status -ne 200) { $FailedDriveLookups++; continue }
                $Site = $SiteByRequestId["$($Response.id)"]
                foreach ($Drive in @($Response.body.value)) {
                    if ($Drive.id) {
                        $DriveEntries.Add([PSCustomObject]@{ Site = $Site; Drive = $Drive })
                    }
                }
            }
        }
        Write-Host "[SharingLinks][$TenantFilter] Phase 2 (site drives): $($DriveEntries.Count) drives, $FailedDriveLookups site drive lookups failed (+$([math]::Round(((Get-Date) - $StartTime).TotalSeconds,1))s)"

        # 3) Delta-scan every drive and keep items that carry the "shared" facet.
        $DriveByRequestId = @{}
        $RequestId = 0
        $DeltaRequests = foreach ($Entry in $DriveEntries) {
            $DriveByRequestId["$RequestId"] = $Entry
            @{
                id     = "$RequestId"
                method = 'GET'
                url    = "drives/$($Entry.Drive.id)/root/delta?`$select=id,name,webUrl,parentReference,file,folder,shared,size,lastModifiedDateTime&`$top=999"
            }
            $RequestId++
        }

        $SharedItems = [System.Collections.Generic.List[object]]::new()
        $FailedDrives = 0
        if (@($DeltaRequests).Count -gt 0) {
            Write-Host "[SharingLinks][$TenantFilter] Phase 3 (delta scan): scanning $(@($DeltaRequests).Count) drives for shared items"
            try {
                $DeltaResponses = New-GraphBulkRequest -tenantid $TenantFilter -Requests @($DeltaRequests) -asapp $true
            } catch {
                Write-Host "[SharingLinks][$TenantFilter] Phase 3 (delta scan) BULK FAILED after $([math]::Round(((Get-Date) - $StartTime).TotalSeconds,1))s: $($_.Exception.Message)"
                throw
            }
            foreach ($Response in $DeltaResponses) {
                if ($Response.status -and $Response.status -ne 200) {
                    $FailedDrives++
                    $FailedEntry = $DriveByRequestId["$($Response.id)"]
                    Write-Host "[SharingLinks][$TenantFilter] Phase 3: drive delta returned status $($Response.status) for '$($FailedEntry.Site.webUrl)' / '$($FailedEntry.Drive.name)' - $($Response.body.error.message)"
                    continue
                }
                $Entry = $DriveByRequestId["$($Response.id)"]
                foreach ($Item in @($Response.body.value)) {
                    if ($Item.shared -and -not $Item.deleted) {
                        $SharedItems.Add([PSCustomObject]@{ Site = $Entry.Site; Drive = $Entry.Drive; Item = $Item })
                    }
                }
            }
        }
        Write-Host "[SharingLinks][$TenantFilter] Phase 3 (delta scan): $($SharedItems.Count) shared items, $FailedDrives drive scans failed (+$([math]::Round(((Get-Date) - $StartTime).TotalSeconds,1))s)"

        # 4) Fetch the permissions of every shared item, in bulk.
        $SharedItemByRequestId = @{}
        $RequestId = 0
        $PermissionRequests = foreach ($Entry in $SharedItems) {
            $SharedItemByRequestId["$RequestId"] = $Entry
            @{
                id     = "$RequestId"
                method = 'GET'
                url    = "drives/$($Entry.Drive.id)/items/$($Entry.Item.id)/permissions"
            }
            $RequestId++
        }

        # 5) Build one row per sharing link (any scope) or direct grant to an external user.
        $Rows = [System.Collections.Generic.List[object]]::new()
        $FailedPermissionLookups = 0
        if (@($PermissionRequests).Count -gt 0) {
            Write-Host "[SharingLinks][$TenantFilter] Phase 4 (permissions): fetching permissions for $(@($PermissionRequests).Count) shared items"
            try {
                $PermissionResponses = New-GraphBulkRequest -tenantid $TenantFilter -Requests @($PermissionRequests) -asapp $true
            } catch {
                Write-Host "[SharingLinks][$TenantFilter] Phase 4 (permissions) BULK FAILED after $([math]::Round(((Get-Date) - $StartTime).TotalSeconds,1))s: $($_.Exception.Message)"
                throw
            }
            foreach ($Response in $PermissionResponses) {
                if ($Response.status -and $Response.status -ne 200) { $FailedPermissionLookups++; continue }
                $Entry = $SharedItemByRequestId["$($Response.id)"]
                $Site = $Entry.Site
                $Drive = $Entry.Drive
                $Item = $Entry.Item

                foreach ($Permission in @($Response.body.value)) {
                    # Only permissions set on the item itself; inherited ones are reported on their parent.
                    if ($Permission.inheritedFrom) { continue }

                    if ($Permission.link) {
                        $Recipients = @($Permission.grantedToIdentitiesV2 ?? $Permission.grantedToIdentities)
                        $LinkScope = $Permission.link.scope ?? 'users'
                        $Classification = switch ($LinkScope) {
                            'anonymous' { 'Anonymous' }
                            'organization' { 'Internal' }
                            'existingAccess' { 'Internal' }
                            default {
                                $HasExternal = $false
                                foreach ($Recipient in $Recipients) {
                                    if (Test-CIPPExternalIdentity -Identity $Recipient -InternalDomains $InternalDomains) { $HasExternal = $true; break }
                                }
                                if ($HasExternal) { 'External' } else { 'Internal' }
                            }
                        }
                        $LinkType = $Permission.link.type ?? 'link'
                        $LinkUrl = $Permission.link.webUrl
                    } else {
                        # Direct grant (no sharing link): only report grants to external users.
                        $Recipients = @($Permission.grantedToV2 ?? $Permission.grantedTo)
                        if ($Permission.roles -contains 'owner') { continue }
                        $HasExternal = $false
                        foreach ($Recipient in $Recipients) {
                            if (Test-CIPPExternalIdentity -Identity $Recipient -InternalDomains $InternalDomains) { $HasExternal = $true; break }
                        }
                        if (-not $HasExternal) { continue }
                        $Classification = 'External'
                        $LinkScope = 'direct'
                        $LinkType = 'directGrant'
                        $LinkUrl = $null
                    }

                    $SharedWith = @($Recipients | ForEach-Object { Get-CIPPIdentityLabel -Identity $_ } | Where-Object { $_ } | Sort-Object -Unique)

                    $Rows.Add([PSCustomObject]@{
                            id                   = "$($Drive.id)_$($Item.id)_$($Permission.id)"
                            siteId               = $Site.id
                            siteName             = $Site.displayName ?? $Site.name
                            siteUrl              = $Site.webUrl
                            workload             = if ($Site.isPersonalSite) { 'OneDrive' } else { 'SharePoint' }
                            driveId              = $Drive.id
                            driveName            = $Drive.name
                            itemId               = $Item.id
                            fileName             = $Item.name
                            itemUrl              = $Item.webUrl
                            itemType             = if ($Item.folder) { 'Folder' } else { 'File' }
                            size                 = $Item.size
                            lastModifiedDateTime = $Item.lastModifiedDateTime
                            permissionId         = $Permission.id
                            linkType             = $LinkType
                            linkScope            = $LinkScope
                            classification       = $Classification
                            roles                = @($Permission.roles)
                            sharedWith           = $SharedWith
                            linkUrl              = $LinkUrl
                            hasPassword          = $Permission.hasPassword ?? $false
                            expirationDateTime   = $Permission.expirationDateTime
                        })
                }
            }
        }

        Write-Host "[SharingLinks][$TenantFilter] Phase 5 (write): writing $($Rows.Count) rows to cache ($FailedPermissionLookups permission lookups failed)"
        Add-CIPPDbItem -TenantFilter $TenantFilter -Type 'SharePointSharingLinks' -Data @($Rows) -AddCount

        $Message = "Cached $($Rows.Count) sharing links across $($DriveEntries.Count) drives" + $(if ($FailedDrives -gt 0) { " ($FailedDrives drives could not be scanned)" } else { '' })
        Write-Host "[SharingLinks][$TenantFilter] DONE: $Message (total $([math]::Round(((Get-Date) - $StartTime).TotalSeconds,1))s)"
        Write-LogMessage -API 'CIPPDBCache' -tenant $TenantFilter -message $Message -sev Debug

    } catch {
        Write-Host "[SharingLinks][$TenantFilter] ABORTED after $([math]::Round(((Get-Date) - $StartTime).TotalSeconds,1))s - existing cache left intact. Error: $($_.Exception.Message)"
        Write-LogMessage -API 'CIPPDBCache' -tenant $TenantFilter -message "Failed to cache SharePoint sharing links: $($_.Exception.Message)" -sev Error -LogData (Get-CippException -Exception $_)
    }
}
