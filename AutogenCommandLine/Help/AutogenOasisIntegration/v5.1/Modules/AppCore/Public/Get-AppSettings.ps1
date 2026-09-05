function Get-AppSettings {
    param (
        [Parameter(Mandatory)]
        [string]$ConfigPath
    )

    # Verify Configuration File Exists
    if (-not (Test-Path $ConfigPath)) {
        throw (New-IntegrationError -Message "Configuration file not found at: $ConfigPath" -ExitCode ([ExitCode]::ConfigNotFound))
    }

    # Import and parse CSV
    try {
        $csvRows = Import-Csv -Path $ConfigPath

        $settings = @{}

        foreach ($row in $csvRows) {
            $group = $row.SettingGroup
            $key   = $row.Key
            $val   = $row.Value

            if (-not $settings.ContainsKey($group)) {
                $settings[$group] = @{}
            }

            $settings[$group][$key] = $val
        }

        return $settings
    }
    catch {
        throw (New-IntegrationError -Message "Failed to parse appsettings CSV. Details: $($_.Exception.Message)" -ExitCode ([ExitCode]::UnknownError))
    }
}
