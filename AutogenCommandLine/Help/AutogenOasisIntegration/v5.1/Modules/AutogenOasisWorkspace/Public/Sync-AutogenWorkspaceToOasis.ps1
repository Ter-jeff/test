function Sync-AutogenWorkspaceToOasis {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$AutogenOutputRoot,

        [Parameter(Mandatory)]
        [string]$ProjectFolder,

        [Parameter(Mandatory)]
        [string]$IgxlProjName
    )

    Write-Log -Level Information -Message "Collecting Autogen output from root: $AutogenOutputRoot"
    $output = Get-AutogenOutput -Root $AutogenOutputRoot

    Write-Log -Level Information -Message "Validating Autogen output structure..."
    Assert-AutogenOutput -Output $output

    Write-Log -Level Information -Message "Checking existing Oasis project for backup..."
    Backup-OasisWorkspace `
        -Output $output `
        -ProjectFolder $ProjectFolder

    Write-Log -Level Information -Message "Applying Autogen output to Oasis project target area..."
    Copy-AutogenToOasis `
        -Output $output `
        -ProjectFolder $ProjectFolder `
        -IgxlProjName $IgxlProjName
}
