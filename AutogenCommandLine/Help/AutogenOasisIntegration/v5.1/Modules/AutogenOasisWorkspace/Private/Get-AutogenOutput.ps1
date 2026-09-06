function Get-AutogenOutput {
    param (
        [Parameter(Mandatory)]
        [string]$Root
    )

    $commonPath = Join-Path $Root 'Common'
    $modulePath = Join-Path $Root 'Module'

    $igxlProj = Get-ChildItem -Path $Root -Filter '*.igxlProj' | Where-Object { -not $_.PSIsContainer } | Select-Object -First 1

    $igxlProjPath = $null
    $igxlProjName = $null

    if ($null -ne $igxlProj) {
        $igxlProjPath = $igxlProj.FullName
        $igxlProjName = $igxlProj.Name
    }

    return [PSCustomObject]@{
        Root = $Root
        CommonPath = $commonPath
        ModulePath = $modulePath
        IgxlProjPath = $igxlProjPath
        IgxlProjName = $igxlProjName
    }
}
