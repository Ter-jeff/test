function Backup-OasisWorkspace {
    param (
        [Parameter(Mandatory)]
        $Output,

        [Parameter(Mandatory)]
        [string]$ProjectFolder
    )

    if (-not (Test-Path $ProjectFolder)) {
        Write-Log -Level Information -Message "Project folder does not exist. Skipping backup."
        return
    }

    # Create timestamped backup target directory
    $timestamp = (Get-Date).ToString('yyyy-MM-dd-HHmmss')
    $backupFolder = Join-Path $ProjectFolder "intg-$timestamp"

    Write-Log -Level Information -Message "Creating backup folder: $backupFolder"
    New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null

    # Backup Common Folder Content
    $sourceCommon = Join-Path $ProjectFolder 'Common'
    if (Test-Path $sourceCommon) {
        Copy-Item -Path $sourceCommon -Destination (Join-Path $backupFolder 'Common') -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Backup Module Folder Content
    $sourceModule = Join-Path $ProjectFolder 'Module'
    if (Test-Path $sourceModule) {
        Copy-Item -Path $sourceModule -Destination (Join-Path $backupFolder 'Module') -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Backup igxlProj File
    if (-not [string]::IsNullOrWhiteSpace($Output.IgxlProjName)) {
        $sourceProj = Join-Path $ProjectFolder $Output.IgxlProjName
        if (Test-Path $sourceProj) {
            Copy-Item -Path $sourceProj -Destination (Join-Path $backupFolder $Output.IgxlProjName) -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Log -Level Information -Message "Backup completed successfully."
}
