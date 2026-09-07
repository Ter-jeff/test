function Invoke-OasisAddinClient {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$CliPath,

        [Parameter(Mandatory)]
        [string]$Params
    )

    if (-not (Test-Path $CliPath -PathType Leaf)) {
        throw (New-IntegrationError `
            -Message "OasisAddInClient executable not found at designated path: $CliPath" `
            -ExitCode ([ExitCode]::UnknownError))
    }

    Start-OasisProcess -CliPath $CliPath -Params $Params
}
