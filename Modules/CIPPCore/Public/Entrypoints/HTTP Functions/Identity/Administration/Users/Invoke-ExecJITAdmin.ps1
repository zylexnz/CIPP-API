function Invoke-ExecJITAdmin {
    <#
    .FUNCTIONALITY
        Entrypoint
    .ROLE
        Identity.Role.ReadWrite

    .DESCRIPTION
        Just-in-time admin management API endpoint. This function can create users, add roles, remove roles, delete, or disable a user.
    #>
    [CmdletBinding()]
    param($Request, $TriggerMetadata)

    $APIName = $Request.Params.CIPPEndpoint
    $Headers = $Request.Headers
    $TenantFilter = $Request.Body.tenantFilter.value ? $Request.Body.tenantFilter.value : $Request.Body.tenantFilter


    if ($Request.Body.existingUser.value -match '^[a-f0-9]{8}-([a-f0-9]{4}-){3}[a-f0-9]{12}$') {
        $Username = (New-GraphGetRequest -uri "https://graph.microsoft.com/v1.0/users/$($Request.Body.existingUser.value)" -tenantid $TenantFilter).userPrincipalName
    }

    $Start = ([System.DateTimeOffset]::FromUnixTimeSeconds($Request.Body.StartDate)).DateTime.ToLocalTime()
    $Expiration = ([System.DateTimeOffset]::FromUnixTimeSeconds($Request.Body.EndDate)).DateTime.ToLocalTime()
    $Results = [System.Collections.Generic.List[string]]::new()

    if ($Request.Body.userAction -eq 'create') {
        $Domain = $Request.Body.Domain.value ? $Request.Body.Domain.value : $Request.Body.Domain
        $Username = "$($Request.Body.Username)@$($Domain)"
        Write-Information "Creating JIT Admin user: $($Request.Body.username)"

        $JITAdmin = @{
            User         = @{
                'FirstName'         = $Request.Body.FirstName
                'LastName'          = $Request.Body.LastName
                'UserPrincipalName' = $Username
            }
            Expiration   = $Expiration
            Reason       = $Request.Body.reason
            Action       = 'Create'
            TenantFilter = $TenantFilter
        }
        $CreateResult = Set-CIPPUserJITAdmin @JITAdmin
        Write-LogMessage -Headers $Headers -API $APIName -tenant $TenantFilter -message "Created JIT Admin user: $Username. Reason: $($Request.Body.reason). Roles: $($Request.Body.adminRoles.label -join ', ')" -Sev 'Info' -LogData $JITAdmin
        $Results.Add("Created User: $Username")
        if (!$Request.Body.UseTAP) {
            $Results.Add("Password: $($CreateResult.password)")
        }
        $Results.Add("JIT Admin Expires: $($Expiration)")
        Start-Sleep -Seconds 1
    }

    #Region TAP creation
    if ($Request.Body.UseTAP) {
        try {
            if ($Start -gt (Get-Date)) {
                $TapParams = @{
                    startDateTime = [System.DateTimeOffset]::FromUnixTimeSeconds($Request.Body.StartDate).DateTime
                }
                $TapBody = ConvertTo-Json -Depth 5 -InputObject $TapParams
            } else {
                $TapBody = '{}'
            }
            # Write-Information "https://graph.microsoft.com/beta/users/$Username/authentication/temporaryAccessPassMethods"
            # Retry creating the TAP up to 10 times, since it can fail due to the user not being fully created yet. Sometimes it takes 2 reties, sometimes it takes 8+. Very annoying. -Bobby
            $Retries = 0
            $MAX_TAP_RETRIES = 10
            do {
                try {
                    $TapRequest = New-GraphPostRequest -uri "https://graph.microsoft.com/beta/users/$($Username)/authentication/temporaryAccessPassMethods" -tenantid $TenantFilter -type POST -body $TapBody
                } catch {
                    Start-Sleep -Seconds 2
                    Write-Information "ERROR: Run $Retries of $MAX_TAP_RETRIES : Failed to create TAP, retrying"
                    # Write-Information ( ConvertTo-Json -Depth 5 -InputObject (Get-CippException -Exception $_))
                }
                $Retries++
            } while ( $null -eq $TapRequest.temporaryAccessPass -and $Retries -le $MAX_TAP_RETRIES )

            $TempPass = $TapRequest.temporaryAccessPass
            $PasswordExpiration = $TapRequest.LifetimeInMinutes

            $PasswordLink = New-PwPushLink -Payload $TempPass
            $Password = $PasswordLink ? $PasswordLink : $TempPass

            $Results.Add("Temporary Access Pass: $Password")
            $Results.Add("This TAP is usable starting at $($TapRequest.startDateTime) UTC for the next $PasswordExpiration minutes")
        } catch {
            $Results.Add('Failed to create TAP, if this is not yet enabled, use the Standards to push the settings to the tenant.')
            Write-Information (Get-CippException -Exception $_ | ConvertTo-Json -Depth 5)
            if ($Password) {
                $Results.Add("Password: $Password")
            }
        }
    }
    #EndRegion TAP creation

    $Parameters = @{
        TenantFilter = $TenantFilter
        User         = @{
            'UserPrincipalName' = $Username
        }
        Roles        = $Request.Body.AdminRoles.value
        Action       = 'AddRoles'
        Reason       = $Request.Body.Reason
        Expiration   = $Expiration
    }
    if ($Start -gt (Get-Date)) {
        $TaskBody = @{
            TenantFilter  = $TenantFilter
            Name          = "JIT Admin (enable): $Username"
            Command       = @{
                value = 'Set-CIPPUserJITAdmin'
                label = 'Set-CIPPUserJITAdmin'
            }
            Parameters    = [pscustomobject]$Parameters
            ScheduledTime = $Request.Body.StartDate
            PostExecution = @{
                Webhook = [bool]($Request.Body.PostExecution | Where-Object -Property value -EQ 'webhook')
                Email   = [bool]($Request.Body.PostExecution | Where-Object -Property value -EQ 'email')
                PSA     = [bool]($Request.Body.PostExecution | Where-Object -Property value -EQ 'PSA')
            }
        }
        Add-CIPPScheduledTask -Task $TaskBody -hidden $false
        if ($Request.Body.userAction -ne 'create') {
            Set-CIPPUserJITAdminProperties -TenantFilter $TenantFilter -UserId $Request.Body.existingUser.value -Expiration $Expiration -Reason $Request.Body.Reason
        }
        $Results.Add("Scheduling JIT Admin enable task for $Username")
        Write-LogMessage -Headers $Headers -API $APIName -message "Scheduling JIT Admin for existing user: $Username. Reason: $($Request.Body.reason). Roles: $($Request.Body.adminRoles.label -join ', ') " -tenant $TenantFilter -Sev 'Info'
    } else {
        $Results.Add("Executing JIT Admin enable task for $Username")
        Set-CIPPUserJITAdmin @Parameters
        Write-LogMessage -Headers $Headers -API $APIName -message "Executing JIT Admin for existing user: $Username. Reason: $($Request.Body.reason). Roles: $($Request.Body.adminRoles.label -join ', ') " -tenant $TenantFilter -Sev 'Info'
    }

    $DisableTaskBody = [pscustomobject]@{
        TenantFilter  = $TenantFilter
        Name          = "JIT Admin ($($Request.Body.ExpireAction.value)): $Username"
        Command       = @{
            value = 'Set-CIPPUserJITAdmin'
            label = 'Set-CIPPUserJITAdmin'
        }
        Parameters    = [pscustomobject]@{
            TenantFilter = $TenantFilter
            User         = @{
                'UserPrincipalName' = $Username
            }
            Roles        = $Request.Body.AdminRoles.value
            Reason       = $Request.Body.Reason
            Action       = $Request.Body.ExpireAction.value
        }
        PostExecution = @{
            Webhook = [bool]($Request.Body.PostExecution | Where-Object -Property value -EQ 'webhook')
            Email   = [bool]($Request.Body.PostExecution | Where-Object -Property value -EQ 'email')
            PSA     = [bool]($Request.Body.PostExecution | Where-Object -Property value -EQ 'PSA')
        }
        ScheduledTime = $Request.Body.EndDate
    }
    $null = Add-CIPPScheduledTask -Task $DisableTaskBody -hidden $false
    $Results.Add("Scheduling JIT Admin $($Request.Body.ExpireAction.value) task for $Username")

    # TODO - We should find a way to have this return a HTTP status code based on the success or failure of the operation. This also doesn't return the results of the operation in a Results hash table, like most of the rest of the API.
    return ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = @{'Results' = @($Results) }
        })

}
