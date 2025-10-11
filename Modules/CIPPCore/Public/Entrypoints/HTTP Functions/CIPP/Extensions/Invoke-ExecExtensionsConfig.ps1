function Invoke-ExecExtensionsConfig {
    <#
    .FUNCTIONALITY
        Entrypoint
    .ROLE
        CIPP.Extension.ReadWrite
    #>
    [CmdletBinding()]
    param($Request, $TriggerMetadata)
    $Headers = $Request.Headers


    $Body = [PSCustomObject]$Request.Body
    $Results = try {
        # Check if NinjaOne URL is set correctly and the instance has at least version 5.6
        if ($Body.NinjaOne.Enabled -eq $true) {
            $AllowedNinjaHostnames = @(
                'app.ninjarmm.com',
                'eu.ninjarmm.com',
                'oc.ninjarmm.com',
                'ca.ninjarmm.com',
                'us2.ninjarmm.com'
            )
            $SetNinjaHostname = $Body.NinjaOne.Instance -replace '/ws', '' -replace 'https://', ''
            if ($AllowedNinjaHostnames -notcontains $SetNinjaHostname) {
                "Error: NinjaOne URL is not allowed. Allowed hostnames are: $($AllowedNinjaHostnames -join ', ')"
            }
        }

        if ($Body.Hudu.NextSync) {
            #parse unixtime for addedtext
            $Timestamp = [datetime]::UnixEpoch.AddSeconds([int]$Body.Hudu.NextSync).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            Register-CIPPExtensionScheduledTasks -Reschedule -NextSync $Body.Hudu.NextSync -Extensions 'Hudu'
            $AddedText = " Next sync will be at $Timestamp."
            $Body.Hudu.NextSync = ''
        }

        $Table = Get-CIPPTable -TableName Extensionsconfig
        foreach ($APIKey in $Body.PSObject.Properties.Name) {
            Write-Information "Working on $apikey"
            if ($Body.$APIKey.APIKey -eq 'SentToKeyVault' -or $Body.$APIKey.APIKey -eq '') {
                Write-Information 'Not sending to keyvault. Key previously set or left blank.'
            } else {
                Write-Information 'writing API Key to keyvault, and clearing.'
                Write-Information "$env:WEBSITE_DEPLOYMENT_ID"
                if ($Body.$APIKey.APIKey) {
                    Set-ExtensionAPIKey -Extension $APIKey -APIKey $Body.$APIKey.APIKey
                }
                if ($Body.$APIKey.PSObject.Properties.Name -notcontains 'APIKey') {
                    $Body.$APIKey | Add-Member -MemberType NoteProperty -Name APIKey -Value 'SentToKeyVault'
                } else {
                    $Body.$APIKey.APIKey = 'SentToKeyVault'
                }
            }
            $Body.$APIKey = $Body.$APIKey | Select-Object * -ExcludeProperty ResetPassword
        }
        $Body = $Body | Select-Object * -ExcludeProperty APIKey, Enabled | ConvertTo-Json -Depth 10 -Compress
        $Config = @{
            'PartitionKey' = 'CippExtensions'
            'RowKey'       = 'Config'
            'config'       = [string]$Body
        }

        Add-CIPPAzDataTableEntity @Table -Entity $Config -Force | Out-Null

        #Write-Information ($Request.Headers | ConvertTo-Json)
        $AddObject = @{
            PartitionKey = 'InstanceProperties'
            RowKey       = 'CIPPURL'
            Value        = [string]([System.Uri]$Headers.'x-ms-original-url').Host
        }
        Write-Information ($AddObject | ConvertTo-Json -Compress)
        $ConfigTable = Get-CIPPTable -tablename 'Config'
        Add-AzDataTableEntity @ConfigTable -Entity $AddObject -Force

        Register-CIPPExtensionScheduledTasks
        "Successfully saved the extension configuration. $AddedText"
    } catch {
        "Failed to save the extensions configuration: $($_.Exception.message) Linenumber: $($_.InvocationInfo.ScriptLineNumber)"
    }



    return ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = @{'Results' = $Results }
        })

}
