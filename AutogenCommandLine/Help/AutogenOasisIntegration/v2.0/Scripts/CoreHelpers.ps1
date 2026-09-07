$Global:ExitCodeMap = @{
    Success = 0
    UnknownError = 1
    LoggerNotInitialized = 10

    ConfigNotFound = 30
    ConfigNoAutogenSection = 31
    ConfigNoAutogenCliPath = 32
    ConfigNoInputInfoPath = 33
    ConfigNoOasisSection = 34
    ConfigNoOasisAddinClientPath = 35
    ConfigNoProjectPath = 36
    InvalidInputInfo = 37
    NoIgxlProjUnderProjectPath = 38
    MultipleIgxlProjsUnderProjectPath = 39

    AutogenFailed = 100
    AutogenNotFound = 101
    AutogenNotExe = 102
    AutogenInputInfoNotFound = 103
    AutogenInputInfoNotCsv = 104
    AutogenOutputInvalid = 105

    OasisAddinClientFailed = 200
    OasisAddinClientConflict = 201
}

function New-IntegrationError {
    param(
        $Message,
        $ExitCodeValue
    )

    $exception = New-Object System.Exception $Message

    Add-Member -InputObject $exception -MemberType NoteProperty -Name 'ExitCode' -Value $ExitCodeValue -Force

    return $exception
}

function Get-AppSettings {
    param($ConfigPath)

    if (-not (Test-Path $ConfigPath)) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message "Config CSV file not found at $ConfigPath" -ExitCodeValue $errCode)
    }

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
        $errCode = $Global:ExitCodeMap.UnknownError
        $failMsg = "Failed to process configuration CSV natively: " + $_.ToString()
        throw (New-IntegrationError -Message $failMsg -ExitCodeValue $errCode)
    }
}

function Get-InputInfo {
    param($CsvPath)

    if (-not (Test-Path $CsvPath)) {
        $errCode = $Global:ExitCodeMap.AutogenInputInfoNotFound
        throw (New-IntegrationError -Message "Input CSV file not found at $CsvPath" -ExitCodeValue $errCode)
    }

    $rawCsv = Import-Csv -Path $CsvPath
    $inputMap = @{}

    foreach ($row in $rawCsv) {
        if ($null -ne $row.Item) {
            $inputMap[$row.Item] = $row
        }
    }

    return $inputMap
}

function Get-OutputFolder {
    param (
        $InputMap
    )

    if (-not $InputMap) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message "Missing mandatory parameter 'InputMap' in Get-OutputFolder." -ExitCodeValue $errCode)
    }

    $outputKey = '--outputFile'

    if (-not $InputMap.ContainsKey($outputKey)) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError `
            -Message "Required property $outputKey was missing in the csv." `
            -ExitCodeValue $errCode)
    }

    $rowObject = $InputMap[$outputKey]
    $rawPath = $null

    if ($null -ne $rowObject) {
        if ($rowObject.PSObject.Properties['value']) {
            $rawPath = $rowObject.value
        }
        elseif ($rowObject.PSObject.Properties['Value']) {
            $rawPath = $rowObject.Value
        }
    }

    if (($null -eq $rawPath) -or ($rawPath.ToString().Trim() -eq "")) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError `
            -Message "The $outputKey property is empty or whitespace inside csv." `
            -ExitCodeValue $errCode)
    }

    $rawPath = $rawPath.ToString().Trim()

    try {
        $absolutePath = [System.IO.Path]::GetFullPath($rawPath)

        return Split-Path -Path $absolutePath -Parent
    }
    catch {
        $errCode = $Global:ExitCodeMap.UnknownError
        $failMsg = "The $outputKey path format is invalid or corrupt: " + $rawPath + ". Details: " + $_.ToString()

        throw (New-IntegrationError -Message $failMsg -ExitCodeValue $errCode)
    }
}

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

    $igxlProj = @(Get-ChildItem -Path $FolderPath -Filter "*.igxlProj" | Where-Object {
        -not $_.PSIsContainer
    })

    $count = @($igxlProj).Count

    if ($count -eq 0) {
        throw (New-IntegrationError `
            -Message "No any '.igxlProj' under  '$FolderPath'" `
            -ExitCode ($Global:ExitCodeMap.NoIgxlProjUnderProjectPath))
    }
    elseif ($count -gt 1) {
        throw (New-IntegrationError `
            -Message "More than one '.igxlProj' under  '$FolderPath'" `
            -ExitCode ($Global:ExitCodeMap.MultipleIgxlProjsUnderProjectPath))
    }

    return $igxlProj[0].Name
}
