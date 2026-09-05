Import-Module AppCore -ErrorAction Stop
Import-Module AppLogging -ErrorAction Stop

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
