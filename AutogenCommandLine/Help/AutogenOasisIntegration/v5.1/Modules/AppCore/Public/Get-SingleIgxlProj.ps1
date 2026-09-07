function Get-SingleIgxlProj {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    if (-not (Test-Path $FolderPath)) {
        throw (New-IntegrationError `
            -Message "Foldre '$FolderPath' not found!" `
            -ExitCode ([ExitCode]::ConfigNoProjectPath))
    }

    $igxlProj = Get-ChildItem -Path $FolderPath -Filter "*.igxlProj" | Where-Object {
        -not $_.PSIsContainer
    }

    $count = @($igxlProj).Count

    if ($count -eq 0) {
        throw (New-IntegrationError `
            -Message "No any '.igxlProj' under  '$FolderPath'" `
            -ExitCode ([ExitCode]::NoIgxlProjUnderProjectPath))
    }
    elseif ($count -gt 1) {
        throw (New-IntegrationError `
            -Message "More than one '.igxlProj' under  '$FolderPath'" `
            -ExitCode ([ExitCode]::MultipleIgxlProjsUnderProjectPath))
    }

    return $igxlProj[0].Name
}
