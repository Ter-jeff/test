function Get-InputInfo {
    param (
        [Parameter(Mandatory)]
        [string]$CsvPath
    )

    # Validate file existence
    if (-not (Test-Path $CsvPath)) {
        throw (New-IntegrationError `
            -Message "InputInfo file not found: $CsvPath" `
            -ExitCode ([ExitCode]::AutogenInputInfoNotFound))
    }

    # Extract and validate content
    $raw = Import-Csv $CsvPath
    if (-not $raw) {
        throw (New-IntegrationError `
            -Message "InputInfo CSV is empty or invalid: $CsvPath" `
            -ExitCode ([ExitCode]::InvalidInputInfo))
    }

    $inputMap = @{}
    foreach ($row in $raw) {
        if ([string]::IsNullOrWhiteSpace($row.Item)) {
            continue
        }

        $canMultipleValue = $null
        if (-not [string]::IsNullOrWhiteSpace($row.CanMultiple)) {
            $cleanBool = $row.CanMultiple.ToString().Trim().ToLower()
            if ($cleanBool -in @('true', '1', 'yes')) {
                $canMultipleValue = $true
            }
            elseif ($cleanBool -in @('false', '0', 'no')) {
                $canMultipleValue = $false
            }
        }

        $record = [PSCustomObject]@{
            Item        = $row.Item
            Value       = $row.Value
            Type        = $row.Type
            CanMultiple = $canMultipleValue
            Description = $row.Description
        }

        $inputMap[$record.Item] = $record
    }

    return $inputMap
}
