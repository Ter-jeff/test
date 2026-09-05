function Copy-AutogenToOasis {
    param (
        [Parameter(Mandatory)]
        $Output,

        [Parameter(Mandatory)]
        [string]$ProjectFolder,

        [Parameter(Mandatory)]
        [string]$IgxlProjName
    )

    if (-not (Test-Path $ProjectFolder)) {
        Write-Log -Level Information -Message "Project folder does not exist. Creating: $ProjectFolder"
        New-Item -ItemType Directory -Path $ProjectFolder -Force | Out-Null
    }

    Sync-MappedOutputFolder -OutputPath $Output.CommonPath -ProjectFolder $ProjectFolder
    Sync-MappedOutputFolder -OutputPath $Output.ModulePath -ProjectFolder $ProjectFolder

    # Copy igxlProj file
    Write-Log -Level Information -Message "Copying IGXL project file: $($Output.IgxlProjName)"
    $sourcePath = Join-Path $Output.Root $Output.igxlProjName
    $destPath = Join-Path $ProjectFolder $Output.igxlProjName
    Copy-Item -Path $sourcePath -Destination $destPath -Force
}
