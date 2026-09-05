function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Verbose','Information','Warning','Error','Fatal')]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message,

        [string]$Scope = ''
    )

    if (-not $script:LogFilePath) {
        throw (New-IntegrationError `
            -Message 'Logger not initialized. Call Initialize-Logger first.' `
            -ExitCode ([ExitCode]::LoggerNotInitialized))
    }

    $currentSeverity = $script:LogLevelSeverity[$Level]
    if ($currentSeverity -lt $script:MinimumLevelValue) {
        return
    }

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $scopeString = if ($Scope) { "[$Scope]" } else { "" }
    $entry = "[$timestamp][$Level]$scopeString $Message"

    $written = $false
    for ($i = 0; $i -lt 3; $i++) {
        try {
            Add-Content -Path $script:LogFilePath -Value $entry -ErrorAction Stop
            $written = $true
            break
        }
        catch {
            Start-Sleep -Milliseconds (10 + (Get-Random -Min 5 -Max 20))
        }
    }

    # Fallback to console directly if the file system is completely locked out
    if (-not $written) {
        Write-Host "[LOGGING WARNING]: Could not write entry to file due to access lock." -ForegroundColor Yellow
    }

    switch ($Level) {
        'Verbose' {
            Write-Host $entry -ForegroundColor DarkGray
        }
        'Information' {
            Write-Host $entry -ForegroundColor Cyan
        }
        'Warning' {
            Write-Host $entry -ForegroundColor Yellow
        }
        'Error' {
            Write-Host $entry -ForegroundColor Magenta
        }
        'Fatal' {
            Write-Host $entry -ForegroundColor White -BackgroundColor DarkRed
        }
    }
}
