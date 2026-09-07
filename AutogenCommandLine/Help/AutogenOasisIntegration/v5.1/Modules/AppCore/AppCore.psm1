enum ExitCode {
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
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        $ExitCode
    )

    $exception = New-Object System.Exception $Message

    Add-Member -InputObject $exception -MemberType NoteProperty -Name 'ExitCode' -Value $ExitCode -Force

    return $exception
}

$publicPath  = Join-Path $PSScriptRoot 'Public'
$privatePath = Join-Path $PSScriptRoot 'Private'

if (Test-Path $publicPath) {
    foreach ($file in Get-ChildItem $publicPath -Filter '*.ps1') {
        . $file.FullName
    }
}

if (Test-Path $privatePath) {
    foreach ($file in Get-ChildItem $privatePath -Filter '*.ps1') {
        . $file.FullName
    }
}
