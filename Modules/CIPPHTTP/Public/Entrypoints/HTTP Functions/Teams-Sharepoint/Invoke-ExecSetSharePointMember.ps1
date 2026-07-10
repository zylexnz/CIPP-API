function Invoke-ExecSetSharePointMember {
    <#
    .FUNCTIONALITY
        Entrypoint
    .ROLE
        Sharepoint.Site.ReadWrite
    .DESCRIPTION
        Adds or removes a member on a SharePoint site. Group-connected sites are managed through
        the backing Microsoft 365 group via Graph; all other site types (communication, classic)
        are managed through the site's associated members group via the SharePoint REST API
        using certificate authentication.
    #>
    [CmdletBinding()]
    param($Request, $TriggerMetadata)

    $APIName = $Request.Params.CIPPEndpoint
    $Headers = $Request.Headers
    $TenantFilter = $Request.Body.tenantFilter
    $UPN = $Request.Body.user.value

    try {
        if ($Request.Body.SharePointType -eq 'Group') {
            if ($Request.Body.GroupID -match '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$') {
                $GroupId = $Request.Body.GroupID
            } else {
                $GroupId = (New-GraphGetRequest -uri "https://graph.microsoft.com/beta/groups?`$filter=mail eq '$($Request.Body.GroupID)' or proxyAddresses/any(x:endsWith(x,'$($Request.Body.GroupID)')) or mailNickname eq '$($Request.Body.GroupID)'" -ComplexFilter -tenantid $TenantFilter).id
            }

            if ($Request.Body.Add -eq $true) {
                $Results = Add-CIPPGroupMember -GroupType 'Team' -GroupID $GroupID -Member $UPN -TenantFilter $TenantFilter -Headers $Headers
            } else {
                $UserID = (New-GraphGetRequest -uri "https://graph.microsoft.com/v1.0/users/$UPN" -tenantid $TenantFilter).id
                $Results = Remove-CIPPGroupMember -GroupType 'Team' -GroupID $GroupID -Member $UserID -TenantFilter $TenantFilter -Headers $Headers
            }
            $StatusCode = [HttpStatusCode]::OK
        } else {
            # Non group-connected site: manage the site's associated members group via the
            # SharePoint REST API with certificate auth.
            $SiteUrl = $Request.Body.URL
            if (-not $SiteUrl) { throw 'No site URL was provided for this site.' }
            if (-not $UPN) { throw 'No user was selected.' }

            $SharePointInfo = Get-SharePointAdminLink -Public $false -tenantFilter $TenantFilter
            $Scope = "$($SharePointInfo.SharePointUrl)/.default"
            $JsonAccept = @{ Accept = 'application/json;odata=nometadata' }
            $BaseUri = "$($SiteUrl.TrimEnd('/'))/_api"

            try {
                $EnsureBody = ConvertTo-Json -Compress -InputObject @{ logonName = "i:0#.f|membership|$UPN" }
                $EnsuredUser = New-GraphPostRequest -uri "$BaseUri/web/ensureuser" -tenantid $TenantFilter -scope $Scope -type POST -body $EnsureBody -contentType 'application/json;odata=nometadata' -AddedHeaders $JsonAccept -UseCertificate -AsApp $true
            } catch {
                throw "Could not resolve $UPN on the site (ensureuser): $($_.Exception.Message)"
            }
            if (-not $EnsuredUser.Id) {
                throw "Could not resolve $UPN on the site."
            }

            if ($Request.Body.Add -eq $true) {
                # Same shape PnP sends: an SP.User entity posted to the group's users
                # collection, which requires the odata=verbose content type.
                $AddBody = ConvertTo-Json -Compress -Depth 5 -InputObject @{
                    '__metadata' = @{ 'type' = 'SP.User' }
                    'LoginName'  = $EnsuredUser.LoginName
                }
                try {
                    $null = New-GraphPostRequest -uri "$BaseUri/web/associatedmembergroup/users" -tenantid $TenantFilter -scope $Scope -type POST -body $AddBody -contentType 'application/json;odata=verbose' -AddedHeaders $JsonAccept -UseCertificate -AsApp $true
                } catch {
                    throw "Could not add $UPN to the site members group: $($_.Exception.Message)"
                }
                $Results = "Successfully added $UPN as a member of $SiteUrl."
            } else {
                try {
                    $null = New-GraphPostRequest -uri "$BaseUri/web/associatedmembergroup/users/removebyid($($EnsuredUser.Id))" -tenantid $TenantFilter -scope $Scope -type POST -body '{}' -contentType 'application/json;odata=nometadata' -AddedHeaders $JsonAccept -UseCertificate -AsApp $true
                } catch {
                    throw "Could not remove $UPN from the site members group: $($_.Exception.Message)"
                }
                $Results = "Successfully removed $UPN as a member of $SiteUrl."
            }
            Write-LogMessage -Headers $Headers -API $APIName -tenant $TenantFilter -message $Results -sev Info
            $StatusCode = [HttpStatusCode]::OK
        }
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        $Results = "Failed to modify members for $($Request.Body.URL ?? $Request.Body.GroupID). Error: $($ErrorMessage.NormalizedError)"
        Write-LogMessage -Headers $Headers -API $APIName -tenant $TenantFilter -message $Results -sev Error -LogData $ErrorMessage
        $StatusCode = [HttpStatusCode]::BadRequest
    }


    return ([HttpResponseContext]@{
            StatusCode = $StatusCode
            Body       = @{ 'Results' = $Results }
        })

}
