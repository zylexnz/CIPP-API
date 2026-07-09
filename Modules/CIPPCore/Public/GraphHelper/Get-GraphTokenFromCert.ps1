function Get-GraphTokenFromCert {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string]$AppId,

        [string]$Scope = 'https://graph.microsoft.com/.default',
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [switch]$SkipCache
    )
    #########################################################
    # Create Bearer Token From Certificate for HBU Graph
    #########################################################

    # Check the token cache before building a new assertion. Uses the shared .NET cache when
    # loaded (cross-runspace), otherwise a script-scope fallback (local dev).
    $UseSharedTokenCache = ($SkipCache -ne $true) -and ($null -ne ('CIPP.CIPPTokenCache' -as [type]))
    if ($UseSharedTokenCache) {
        $TokenCacheKey = [CIPP.CIPPTokenCache]::BuildKey([string]$TenantId, [string]$Scope, $true, [string]$AppId, 'client_credentials_certificate')
        $CacheEntry = [CIPP.CIPPTokenCache]::Lookup($TokenCacheKey, 120)
        if ($CacheEntry.Found -and -not [string]::IsNullOrWhiteSpace($CacheEntry.TokenPayloadJson)) {
            try {
                return ($CacheEntry.TokenPayloadJson | ConvertFrom-Json -ErrorAction Stop)
            } catch {
                [CIPP.CIPPTokenCache]::Remove($TokenCacheKey)
            }
        }
    } elseif ($SkipCache -ne $true) {
        $ScriptCacheKey = "$TenantId|$AppId|$Scope"
        $Cached = $script:CertTokenCache.$ScriptCacheKey
        if ($Cached.expires_on -and [int](Get-Date -UFormat %s -Millisecond 0) -lt ($Cached.expires_on - 120)) {
            return $Cached
        }
    }

    # get sha256 hash of certificate
    $sha256 = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    $hash = $sha256.ComputeHash($Certificate.RawData)
    $hash = [Convert]::ToBase64String($hash)

    # Create JWT timestamp for expiration
    $StartDate = (Get-Date '1970-01-01T00:00:00Z' ).ToUniversalTime()
    $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
    $JWTExpiration = [math]::Round($JWTExpirationTimeSpan, 0)

    # Create JWT validity start timestamp
    $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
    $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan, 0)

    # Create JWT header
    $JWTHeader = @{
        alg        = 'PS256'
        typ        = 'JWT'
        # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
        'x5t#S256' = $hash -replace '\+', '-' -replace '/', '_' -replace '='
    }

    # Create JWT payload
    $JWTPayLoad = @{
        # Issuer = your application
        iss = $AppId

        # What endpoint is allowed to use this JWT
        aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

        # JWT ID: random guid
        jti = [guid]::NewGuid()

        # Expiration timestamp
        exp = $JWTExpiration

        # Not to be used before
        nbf = $NotBefore

        # JWT Subject
        sub = $AppId
    }

    # Convert header and payload to base64
    $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
    $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

    $JWTPayLoadToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
    $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

    # Join header and Payload with "." to create a valid (unsigned) JWT
    $JWT = $EncodedHeader + '.' + $EncodedPayload

    # Get the private key object of your certificate
    # $PrivateKey = $Certificate.PrivateKey
    $PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))

    # Define RSA signature and hashing algorithm
    $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pss
    $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

    # Create a signature of the JWT
    $Signature = [Convert]::ToBase64String(
        $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT), $HashAlgorithm, $RSAPadding)
    ) -replace '\+', '-' -replace '/', '_' -replace '='

    # Join the signature to the JWT with "."
    $JWT = $JWT + '.' + $Signature

    # Create a hash with body parameters
    $Body = @{
        client_id             = $AppId
        client_assertion      = $JWT
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        scope                 = $Scope
        grant_type            = 'client_credentials'
    }

    $Url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    # Use the self-generated JWT as Authorization
    $Header = @{
        Authorization = "Bearer $JWT"
    }

    # Splat the parameters for Invoke-Restmethod for cleaner code
    $PostSplat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = $Body
        Uri         = $Url
        Headers     = $Header
    }

    # AADSTS700027 occurs transiently after certificate rotation while the load-balanced
    # token service nodes catch up with the directory - retry briefly before giving up.
    $MaxRetries = 3
    for ($Attempt = 1; $Attempt -le $MaxRetries; $Attempt++) {
        try {
            $AccessToken = Invoke-CIPPRestMethod @PostSplat -ErrorAction Stop
            if ($AccessToken.access_token) {
                if ($null -eq $AccessToken.expires_on -and $AccessToken.expires_in) {
                    $ExpiresOn = [int](Get-Date -UFormat %s -Millisecond 0) + $AccessToken.expires_in
                    Add-Member -InputObject $AccessToken -NotePropertyName 'expires_on' -NotePropertyValue $ExpiresOn -Force
                }
                if ($UseSharedTokenCache) {
                    try {
                        [CIPP.CIPPTokenCache]::Store($TokenCacheKey, ($AccessToken | ConvertTo-Json -Depth 20 -Compress), [int64]$AccessToken.expires_on)
                    } catch {
                        # Ignore shared cache write failures
                    }
                } elseif ($SkipCache -ne $true) {
                    if (-not $script:CertTokenCache) { $script:CertTokenCache = [HashTable]::Synchronized(@{}) }
                    $script:CertTokenCache.$ScriptCacheKey = $AccessToken
                }
            }
            return $AccessToken
        } catch {
            if ($Attempt -lt $MaxRetries -and $_.ErrorDetails.Message -match 'AADSTS700027') {
                Write-Warning "Certificate not yet recognized by the token service (attempt $Attempt of $MaxRetries). Retrying in 10 seconds."
                Start-Sleep -Seconds 10
            } else {
                Write-Error $_
                return
            }
        }
    }

}
