function Sync-MappedOutputFolder {
    param (
        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter(Mandatory)]
        [string]$ProjectFolder
    )
    $folderName = Split-Path $OutputPath -Leaf

    # Copy folder contents safely
    Write-Log -Level Information -Message "Copying $folderName folder..."
    $targetPath = Join-Path $ProjectFolder $folderName
    if (Test-Path $targetPath) {
        Remove-Item -Path $targetPath -Recurse -Force
    }
    Copy-Item -Path $OutputPath -Destination $ProjectFolder -Recurse -Force
}
