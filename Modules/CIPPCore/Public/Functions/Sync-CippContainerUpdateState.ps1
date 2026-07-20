function Sync-CippContainerUpdateState {
    <#
    .SYNOPSIS
    Reconcile the stored container update-check result with the build that is actually running.

    .DESCRIPTION
    The ContainerUpdateSettings table is only written when an update check runs, so after a
    restart that applied an update the table keeps reporting the previous build's state —
    "update available" with the old running version — until the next check. This compares the
    stored result against the running APP_VERSION/IMAGE_TAG and clears or recomputes the flag
    locally, without a registry call. Returns the current settings entity, or $null when no
    settings have been saved yet.
    #>
    [CmdletBinding()]
    param()

    $SettingsTable = Get-CippTable -tablename 'ContainerUpdateSettings'
    $Settings = Get-CIPPAzDataTableEntity @SettingsTable -Filter "PartitionKey eq 'Settings' and RowKey eq 'UpdateConfig'" | Select-Object -First 1
    if (-not $Settings) { return $null }

    $RunningVersion = $env:APP_VERSION
    $StoredRunning = [string]($Settings.RunningVersion ?? '')
    if (-not $RunningVersion -or -not $StoredRunning -or $StoredRunning -eq $RunningVersion) {
        return $Settings
    }

    # The container restarted onto a different build since the last check.
    $StoredRemote = [string]($Settings.RemoteVersion ?? '')
    $CheckedTag = [string]($Settings.CheckedTag ?? '')
    $RunningTag = $env:IMAGE_TAG
    if ($CheckedTag -and $RunningTag -and $CheckedTag -ne $RunningTag) {
        # The last check ran against a different channel tag — its result no longer applies.
        $UpdateAvailable = $false
    } else {
        $UpdateAvailable = [bool]($StoredRemote -and $StoredRemote -ne $RunningVersion)
    }

    $Entity = @{
        PartitionKey    = 'Settings'
        RowKey          = 'UpdateConfig'
        AutoUpdate      = [string]($Settings.AutoUpdate ?? 'false')
        CheckInterval   = [string]($Settings.CheckInterval ?? '0')
        CheckTime       = [string]($Settings.CheckTime ?? '')
        LastCheck       = [string]($Settings.LastCheck ?? '')
        UpdateAvailable = [string]$UpdateAvailable
        RunningVersion  = [string]$RunningVersion
        RemoteVersion   = $StoredRemote
        RemoteDigest    = [string]($Settings.RemoteDigest ?? '')
        RemoteBuildDate = [string]($Settings.RemoteBuildDate ?? '')
        CheckedTag      = $CheckedTag
    }
    Add-CIPPAzDataTableEntity @SettingsTable -Entity $Entity -Force | Out-Null
    Write-Information "Container update state reconciled: running build changed '$StoredRunning' -> '$RunningVersion', UpdateAvailable=$UpdateAvailable"
    return [pscustomobject]$Entity
}
