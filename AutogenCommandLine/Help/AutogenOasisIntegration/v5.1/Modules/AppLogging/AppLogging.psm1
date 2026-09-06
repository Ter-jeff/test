Import-Module AppCore -ErrorAction Stop

$script:LogFilePath = $null

$script:LogLevelSeverity = [ordered]@{
    'Verbose'     = 1
    'Information' = 2
    'Warning'     = 3
    'Error'       = 4
    'Fatal'       = 5
}
$script:MinimumLevelValue = 2

function Initialize-Logger {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$LogDirectory,

        [ValidateSet('Verbose','Information','Warning','Error','Fatal')]
        [string]$MinimumLevel = 'Information'
    )

    if (-not [System.IO.Path]::IsPathRooted($LogDirectory)) {
        $LogDirectory = Join-Path $PSScriptRoot $LogDirectory
    }

    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $timestamp = (Get-Date).ToString('yyyyMMddHHmmss')
    $logFileName = "log-51-$timestamp.log"

    $script:LogFilePath = Join-Path $LogDirectory $logFileName

    $script:MinimumLevelValue = $script:LogLevelSeverity[$MinimumLevel]
}

$publicPath = Join-Path $PSScriptRoot 'Public'
if (Test-Path $publicPath) {
    foreach ($file in Get-ChildItem $publicPath -Filter '*.ps1') {
        . $file.FullName
    }
}
