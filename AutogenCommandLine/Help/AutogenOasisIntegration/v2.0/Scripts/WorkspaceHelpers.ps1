function Get-AutogenOutput {
    param (
        $Root
    )

    if (-not $Root) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message "Missing mandatory parameter 'Root' in Get-AutogenOutput." -ExitCodeValue $errCode)
    }

    $commonPath = Join-Path $Root 'Common'
    $modulePath = Join-Path $Root 'Module'

    $igxlProj = Get-ChildItem -Path $Root -Filter '*.igxlProj' | Where-Object { -not $_.PSIsContainer } | Select-Object -First 1

    $igxlProjPath = $null
    $igxlProjName = $null

    if ($null -ne $igxlProj) {
        $igxlProjPath = $igxlProj.FullName
        $igxlProjName = $igxlProj.Name
    }

    $output = New-Object PSObject
    Add-Member -InputObject $output -MemberType NoteProperty -Name "Root" -Value $Root
    Add-Member -InputObject $output -MemberType NoteProperty -Name "CommonPath" -Value $commonPath
    Add-Member -InputObject $output -MemberType NoteProperty -Name "ModulePath" -Value $modulePath
    Add-Member -InputObject $output -MemberType NoteProperty -Name "IgxlProjPath" -Value $igxlProjPath
    Add-Member -InputObject $output -MemberType NoteProperty -Name "IgxlProjName" -Value $igxlProjName

    return $output
}


function Assert-AutogenOutput {
    param (
        $Output
    )

    if (-not $Output) {
        $errCode = $Global:ExitCodeMap.AutogenOutputInvalid
        throw (New-IntegrationError -Message "Missing mandatory parameter 'Output' in Assert-AutogenOutput." -ExitCodeValue $errCode)
    }

    if (-not (Test-Path $Output.CommonPath)) {
        $errCode = $Global:ExitCodeMap.AutogenOutputInvalid
        $msg = "Autogen output missing Common folder: " + $Output.CommonPath
        throw (New-IntegrationError -Message $msg -ExitCodeValue $errCode)
    }

    if (-not (Test-Path $Output.ModulePath)) {
        $errCode = $Global:ExitCodeMap.AutogenOutputInvalid
        $msg = "Autogen output missing Module folder: " + $Output.ModulePath
        throw (New-IntegrationError -Message $msg -ExitCodeValue $errCode)
    }

    $isProjPathInvalid = (-not $Output.IgxlProjPath) -or ($Output.IgxlProjPath.Trim() -eq "")
    if ($isProjPathInvalid -or (-not (Test-Path $Output.IgxlProjPath))) {
        $errCode = $Global:ExitCodeMap.AutogenOutputInvalid
        $msg = "Autogen output missing .igxlProj file: " + $Output.IgxlProjPath
        throw (New-IntegrationError -Message $msg -ExitCodeValue $errCode)
    }
}


function Backup-OasisWorkspace {
    param (
        $Output,
        $ProjectFolder
    )

    if (-not $Output -or -not $ProjectFolder) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message "Missing mandatory parameters in Backup-OasisWorkspace." -ExitCodeValue $errCode)
    }

    if (-not (Test-Path $ProjectFolder)) {
        Write-Log -Level 'Information' -Message "Project folder does not exist. Skipping backup."
        return
    }

    $timestamp = [DateTime]::Now.ToString('yyyy-MM-dd-HHmmss')
    $backupFolder = Join-Path $ProjectFolder ("intg-" + $timestamp)

    Write-Log -Level 'Information' -Message ("Creating backup folder: " + $backupFolder)
    $null = New-Item -ItemType "Directory" -Path $backupFolder -Force

    $sourceCommon = Join-Path $ProjectFolder 'Common'
    if (Test-Path $sourceCommon) {
        Copy-Item -Path $sourceCommon -Destination (Join-Path $backupFolder 'Common') -Recurse -Force -ErrorAction SilentlyContinue
    }

    $sourceModule = Join-Path $ProjectFolder 'Module'
    if (Test-Path $sourceModule) {
        Copy-Item -Path $sourceModule -Destination (Join-Path $backupFolder 'Module') -Recurse -Force -ErrorAction SilentlyContinue
    }

    $hasProjName = ($null -ne $Output.IgxlProjName) -and ($Output.IgxlProjName.Trim() -ne "")
    if ($hasProjName) {
        $sourceProj = Join-Path $ProjectFolder $Output.IgxlProjName
        if (Test-Path $sourceProj) {
            Copy-Item -Path $sourceProj -Destination (Join-Path $backupFolder $Output.IgxlProjName) -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Log -Level 'Information' -Message "Backup completed successfully."
}

function Copy-AutogenToOasis {
    param (
        $Output,
        $ProjectFolder,
        $IgxlProjName
    )

    if (-not $Output -or -not $ProjectFolder) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message "Missing mandatory parameters ProjectFolder in Copy-AutogenToOasis." -ExitCodeValue $errCode)
    }

    if (-not (Test-Path $ProjectFolder)) {
        Write-Log -Level 'Information' -Message ("Project folder does not exist. Creating: " + $ProjectFolder)
        $null = New-Item -ItemType "Directory" -Path $ProjectFolder -Force
    }

    Sync-MappedOutputFolder -OutputPath $Output.CommonPath -ProjectFolder $ProjectFolder
    Sync-MappedOutputFolder -OutputPath $Output.ModulePath -ProjectFolder $ProjectFolder

    Write-Log -Level 'Information' -Message ("Copying IGXL project file: " + $Output.IgxlProjName)
    $sourcePath = Join-Path $Output.Root $Output.igxlProjName
    $destPath = Join-Path $ProjectFolder $Output.igxlProjName
    Copy-Item -Path $sourcePath -Destination $destPath -Force
}

function Sync-AutogenWorkspaceToOasis {
    param (
        $AutogenOutputRoot,
        $ProjectFolder,
        $IgxlProjName
    )

    if (-not $AutogenOutputRoot -or -not $ProjectFolder) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message "Missing mandatory parameters in Sync-AutogenWorkspaceToOasis." -ExitCodeValue $errCode)
    }

    Write-Log -Level 'Information' -Message ("Collecting Autogen output from root: " + $AutogenOutputRoot)
    $output = Get-AutogenOutput -Root $AutogenOutputRoot

    Write-Log -Level 'Information' -Message "Validating Autogen output structure..."
    Assert-AutogenOutput -Output $output

    Write-Log -Level 'Information' -Message "Checking existing Oasis project for backup..."
    Backup-OasisWorkspace -Output $output -ProjectFolder $ProjectFolder

    Write-Log -Level 'Information' -Message "Applying Autogen output to Oasis project target area..."
    Copy-AutogenToOasis -Output $output -ProjectFolder $ProjectFolder -IgxlProjName $IgxlProjName
}

function Sync-MappedOutputFolder {
    param (
        $OutputPath,
        $ProjectFolder
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
