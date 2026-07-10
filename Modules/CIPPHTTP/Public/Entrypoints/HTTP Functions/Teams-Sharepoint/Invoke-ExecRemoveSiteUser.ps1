function Invoke-ExecRemoveSiteUser {
    <#
    .FUNCTIONALITY
        Entrypoint
    .ROLE
        Sharepoint.Site.ReadWrite
    .DESCRIPTION
        Removes a user from an entire SharePoint site: deleting the site user removes them
        from every site group and direct permission grant at once. Uses the SharePoint REST
        API with certificate authentication. Note this does not revoke sharing links the user
        received by mail; use the sharing link actions for those.
    #>
    [CmdletBinding()]
    param($Request, $TriggerMetadata)

    $APIName = $Request.Params.CIPPEndpoint
    $Headers = $Request.Headers
    $TenantFilter = $Request.Body.tenantFilter
    $SiteUrl = $Request.Body.SiteUrl
    # The picker supplies the SP claims login via addedFields; fall back to building it from the UPN.
    $LoginName = $Request.Body.user.addedFields.LoginName ?? $Request.Body.user.value
    $Label = $Request.Body.user.value ?? $LoginName

    try {
        if (-not $SiteUrl) { throw 'SiteUrl is required.' }
        if (-not $LoginName) { throw 'No user was selected.' }
        if ($LoginName -notmatch '\|') { $LoginName = "i:0#.f|membership|$LoginName" }

        $SharePointInfo = Get-SharePointAdminLink -Public $false -tenantFilter $TenantFilter
        $Scope = "$($SharePointInfo.SharePointUrl)/.default"
        $JsonAccept = @{ Accept = 'application/json;odata=nometadata' }
        $BaseUri = "$($SiteUrl.TrimEnd('/'))/_api"

        try {
            $EnsureBody = ConvertTo-Json -Compress -InputObject @{ logonName = $LoginName }
            $EnsuredUser = New-GraphPostRequest -uri "$BaseUri/web/ensureuser" -tenantid $TenantFilter -scope $Scope -type POST -body $EnsureBody -contentType 'application/json;odata=nometadata' -AddedHeaders $JsonAccept -UseCertificate -AsApp $true
        } catch {
            throw "Could not resolve $Label on the site (ensureuser): $($_.Exception.Message)"
        }
        if (-not $EnsuredUser.Id) { throw "Could not resolve $Label on the site." }
        if ($EnsuredUser.IsSiteAdmin) {
            throw "$Label is a site collection admin. Remove their admin permission first (Remove Site Admin action)."
        }

        try {
            $null = New-GraphPostRequest -uri "$BaseUri/web/siteusers/removebyid($($EnsuredUser.Id))" -tenantid $TenantFilter -scope $Scope -type POST -body '{}' -contentType 'application/json;odata=nometadata' -AddedHeaders $JsonAccept -UseCertificate -AsApp $true
        } catch {
            throw "Could not remove $Label from the site: $($_.Exception.Message)"
        }

        $Results = "Successfully removed $Label from $SiteUrl (all site groups and direct permissions)."
        Write-LogMessage -Headers $Headers -API $APIName -tenant $TenantFilter -message $Results -sev Info
        $StatusCode = [HttpStatusCode]::OK
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        $Results = "Failed to remove $Label from $($SiteUrl): $($ErrorMessage.NormalizedError)"
        Write-LogMessage -Headers $Headers -API $APIName -tenant $TenantFilter -message $Results -sev Error -LogData $ErrorMessage
        $StatusCode = [HttpStatusCode]::BadRequest
    }

    return ([HttpResponseContext]@{
            StatusCode = $StatusCode
            Body       = @{ 'Results' = $Results }
        })
}
