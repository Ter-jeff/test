param(
    $Environment = 'Production',
    $Mode = 'All'
)

# Safely calculate the script root directory without using $PSScriptRoot
$RealScriptRoot = $PSScriptRoot
if (-not $RealScriptRoot) {
    $RealScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$psVersion = $PSVersionTable.PSVersion
$major     = $psVersion.Major
$minor     = $psVersion.Minor

# 5.1 is far from modern, but at least for Teradyne device, it is very modern
$isModernEnvironment = ($major -gt 5) -or ($major -eq 5 -and $minor -ge 1)

if ($isModernEnvironment) {
    # MODERN POWERSHELL EXECUTION PATH (v5.1 / v7+)
    $v5Root = Join-Path $RealScriptRoot "v5.1"
    # Load our modules into local modules path. This can make import much easier.
    $localModulesPath = Join-Path $v5Root "Modules"
    if ($env:PSModulePath -notlike "*$localModulesPath*") {
        $env:PSModulePath = "$localModulesPath;${env:PSModulePath}"
    }

    # Pass explicitly to modern environment
    . (Join-Path $v5Root "AutogenOasisIntegration.ps1") -Environment $Environment -Mode $Mode
}
else {
    # LEGACY FALLBACK PATH (v2.0)
    $v2Root = Join-Path $RealScriptRoot "v2.0"

    # Pass explicitly to legacy environment
    . (Join-Path $v2Root "AutogenOasisIntegration20.ps1") -Environment $Environment -Mode $Mode
}
