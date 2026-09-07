function Invoke-Autogen {
    param (
        [Parameter(Mandatory)]
        [string]$CliPath,

        [Parameter(Mandatory)]
        [string]$InputInfoPath
    )

    if (-not (Test-Path $CliPath -PathType Leaf)) {
        throw (New-IntegrationError `
            -Message "Autogen CLI executable not found: $CliPath" `
            -ExitCode ([ExitCode]::AutogenNotFound))
    }

    if ([System.IO.Path]::GetExtension($CliPath).ToLower() -ne '.exe') {
        throw (New-IntegrationError `
            -Message "Autogen CLI path is not an executable (.exe): $CliPath" `
            -ExitCode ([ExitCode]::AutogenNotExe))
    }

    if (-not (Test-Path $InputInfoPath -PathType Leaf)) {
        throw (New-IntegrationError `
            -Message "inputinfo.csv not found: $InputInfoPath" `
            -ExitCode ([ExitCode]::AutogenInputInfoNotFound))
    }

    if ([System.IO.Path]::GetExtension($InputInfoPath).ToLower() -ne '.csv') {
        throw (New-IntegrationError `
            -Message "inputinfo path is not a spreadsheet file (.csv): $InputInfoPath" `
            -ExitCode ([ExitCode]::AutogenInputInfoNotCsv))
    }

    Start-AutogenProcess -CliPath $CliPath -InputInfoPath $InputInfoPath
}
