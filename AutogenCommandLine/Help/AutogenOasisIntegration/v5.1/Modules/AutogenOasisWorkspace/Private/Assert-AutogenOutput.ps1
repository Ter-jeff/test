function Assert-AutogenOutput {
    param (
        [Parameter(Mandatory)]
        $Output
    )

    if (-not (Test-Path $Output.CommonPath)) {
        throw (New-IntegrationError `
            -Message "Autogen output missing Common folder: $($Output.CommonPath)" `
            -ExitCode ([ExitCode]::AutogenOutputInvalid))
    }

    if (-not (Test-Path $Output.ModulePath)) {
        throw (New-IntegrationError `
            -Message "Autogen output missing Module folder: $($Output.ModulePath)" `
            -ExitCode ([ExitCode]::AutogenOutputInvalid))
    }

    if ([string]::IsNullOrWhiteSpace($Output.IgxlProjPath) -or (-not (Test-Path $Output.IgxlProjPath))) {
        throw (New-IntegrationError `
            -Message "Autogen output missing .igxlProj file: $($Output.IgxlProjPath)" `
            -ExitCode ([ExitCode]::AutogenOutputInvalid))
    }
}
