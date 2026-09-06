function Get-OutputFolder {
    param (
        [Parameter(Mandatory)]
        [hashtable]$InputMap
    )

    $outputKey = '--outputFile'

    # Validate that the key exists in the map
    if (-not $InputMap.ContainsKey($outputKey)) {
        throw (New-IntegrationError `
            -Message "Required property $outputKey was missing in the csv." `
            -ExitCode ([ExitCode]::InvalidInputInfo))
    }

    # Extract and sanitize the raw string value
    $rawPath = $InputMap[$outputKey].Value
    if ([string]::IsNullOrWhiteSpace($rawPath)) {
        throw (New-IntegrationError `
            -Message "The $outputKey property is empty or whitespace inside csv." `
            -ExitCode ([ExitCode]::InvalidInputInfo))
    }

    # Trim accidental spaces left inside the CSV cell
    $rawPath = $rawPath.Trim()

    try {
        $absolutePath = [System.IO.Path]::GetFullPath($rawPath)

        return Split-Path -Path $absolutePath -Parent
    }
    catch {
        throw (New-IntegrationError `
            -Message "The 'outputFile' path format is invalid or corrupt: $rawPath. Details: $($_.Exception.Message)" `
            -ExitCode ([ExitCode]::InvalidInputInfo))
    }
}
