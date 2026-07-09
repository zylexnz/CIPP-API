
function Get-CIPPAuthentication {
    [CmdletBinding()]
    param (
        $APIName = 'Get Keyvault Authentication',
        [switch]$Force
    )
    $Variables = @('ApplicationID', 'ApplicationSecret', 'TenantID', 'RefreshToken')

    try {
        if ($env:AzureWebJobsStorage -eq 'UseDevelopmentStorage=true' -or $env:NonLocalHostAzurite -eq 'true') {
            $Table = Get-CIPPTable -tablename 'DevSecrets'
            $Secret = Get-AzDataTableEntity @Table -Filter "PartitionKey eq 'Secret' and RowKey eq 'Secret'"
            if (!$Secret) {
                throw 'Development variables not set'
            }
            foreach ($Var in $Variables) {
                if ($Secret.$Var) {
                    Set-Item -Path env:$Var -Value $Secret.$Var -Force -ErrorAction Stop
                }
            }
            Write-Host "Got secrets from dev storage. ApplicationID: $env:ApplicationID"
        } else {
            $keyvaultname = Get-CippKeyVaultName
            $Variables | ForEach-Object {
                Set-Item -Path env:$_ -Value (Get-CippKeyVaultSecret -VaultName $keyvaultname -Name $_ -AsPlainText -ErrorAction Stop) -Force
            }
        }
        # Preload the SAM certificate PFX alongside the other credentials. Non-fatal: the
        # certificate does not exist until first bootstrap (setup wizard or weekly timer).
        try {
            if ($env:AzureWebJobsStorage -eq 'UseDevelopmentStorage=true' -or $env:NonLocalHostAzurite -eq 'true') {
                if ($Secret.SAMCertificate) {
                    $env:SAMCertificate = $Secret.SAMCertificate
                }
            } else {
                $SAMCertificate = Get-CippKeyVaultSecret -VaultName $keyvaultname -Name 'SAMCertificate' -AsPlainText -ErrorAction Stop
                if ($SAMCertificate) {
                    $env:SAMCertificate = $SAMCertificate
                }
            }
        } catch {
            Write-Information "SAM certificate not preloaded (not created yet?): $($_.Exception.Message)"
        }

        $env:SetFromProfile = $true
        Write-LogMessage -message 'Reloaded authentication data from KeyVault' -Sev 'debug' -API 'CIPP Authentication'

        return $true
    } catch {
        Write-LogMessage -message 'Could not retrieve keys from Keyvault' -Sev 'CRITICAL' -API 'CIPP Authentication' -LogData (Get-CippException -Exception $_)
        return $false
    }
}
