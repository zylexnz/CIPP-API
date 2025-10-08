function Set-CIPPUserJITAdminProperties {
    [CmdletBinding()]
    param(
        [string]$TenantFilter,
        [string]$UserId,
        [switch]$Enabled,
        $Expiration,
        [switch]$Clear,
        [string]$Reason
    )
    try {
        $Schema = Get-CIPPSchemaExtensions | Where-Object { $_.id -match '_cippUser' } | Select-Object -First 1
        if ($Clear.IsPresent) {
            $Body = [PSCustomObject]@{
                "$($Schema.id)" = @{
                    jitAdminEnabled    = $null
                    jitAdminExpiration = $null
                    jitAdminReason     = $null
                }
            }
        } else {
            $Body = [PSCustomObject]@{
                "$($Schema.id)" = @{
                    jitAdminEnabled    = $Enabled.IsPresent
                    jitAdminExpiration = $Expiration.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                    jitAdminReason     = $Reason
                }
            }
        }

        $Json = ConvertTo-Json -Depth 5 -InputObject $Body
        Write-Information $Json
        New-GraphPOSTRequest -type PATCH -Uri "https://graph.microsoft.com/beta/users/$UserId" -Body $Json -tenantid $TenantFilter | Out-Null
    } catch {
        Write-Information "Error setting JIT Admin properties: $($_.Exception.Message) - $($_.InvocationInfo.PositionMessage)"
    }
}
