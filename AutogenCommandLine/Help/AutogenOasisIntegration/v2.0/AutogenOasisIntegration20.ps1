param(
    $Environment = 'Production',
    $Mode = 'All'
)

$RealScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptDir = Join-Path $RealScriptRoot "Scripts"

# Dot-source our scripts
. (Join-Path $scriptDir "CoreHelpers.ps1")
. (Join-Path $scriptDir "LoggingHelpers.ps1")
. (Join-Path $scriptDir "AutogenWorker.ps1")
. (Join-Path $scriptDir "WorkspaceHelpers.ps1")
. (Join-Path $scriptDir "OasisWorker.ps1")

function Invoke-Main {
    try {
        $repoRoot = Split-Path -Parent $RealScriptRoot

        # Initialize our logging stream inside the central .logs directory
        $logDirectory = Join-Path $repoRoot ".logs"
        Initialize-Logger -LogDirectory $logDirectory

        Write-Log -Level 'Information' -Message "[SYSTEM] Running Powershell 2.0 in '$Environment' environment, '$Mode' mode"

        $configName = "appsettings.csv"
        if ($Environment -eq 'Development') { $configName = "appsettings.Development.csv" }
        $configPath = Join-Path $repoRoot $configName
        Write-Log -Level 'Information' -Message "Using config file: $configPath"

        Start-Integration -ConfigPath $configPath
        exit $Global:ExitCodeMap.Success
    }
    catch {
        $errorRecord = $_

        if ($null -ne $errorRecord.Exception -and $errorRecord.Exception.PSObject.Properties['ExitCode']) {
            # Safely logs expected failures built via New-IntegrationError
            Write-ExpectedError -ErrorRecord $errorRecord

            $exitCodeValue = $errorRecord.Exception.ExitCode
            if ($null -ne $exitCodeValue.value__) { $exitCodeValue = $exitCodeValue.value__ }
            exit $exitCodeValue
        }
        else {
            # Catch-all runtime parsing failure logger
            Write-UnexpectedError -ErrorRecord $errorRecord
            exit $Global:ExitCodeMap.UnknownError
        }
    }
}

function Start-Integration {
    param($ConfigPath)

    Write-Log -Level 'Information' -Message "[START] Initializing configuration..."
    $appSettings = Get-AppSettings -ConfigPath $ConfigPath

    $autogenConfig = $appSettings["autogen"]
    $cliPath       = $autogenConfig["cliPath"]
    $inputInfoPath = $autogenConfig["inputInfoPath"]

    $oasisConfig     = $appSettings["oasis"]
    $addinClientPath = $oasisConfig["addinClientPath"]
    $projectFolder   = $oasisConfig["projectFolder"]

    Write-Log -Level 'Information' -Message "AutogenCommandLine path: $cliPath"
    Write-Log -Level 'Information' -Message "InputInfo path: $inputInfoPath"
    Write-Log -Level 'Information' -Message "OasisAddinClient path: $addinClientPath"
    Write-Log -Level 'Information' -Message "OasisAddinClient projectFolder: $projectFolder"
    $inputInfoMap = Get-InputInfo -CsvPath $inputInfoPath
    $outputFolder = Get-OutputFolder -InputMap $inputInfoMap

    $autogenOutputRootPath = $outputFolder + "\IGLink"
    $oasisProjectFolderPath  = $appSettings["oasis"]["projectFolder"]

    Write-Log -Level 'Information' -Message "Output folder: $outputFolder"
    Write-Log -Level 'Information' -Message "[DONE] Configuration initialized!"

    switch ($Mode) {
        'All' {
            $igxlProjName = Get-SingleIgxlProj -FolderPath $projectFolder
            Write-Log -Level 'Information' -Message "Target igxlProj: $igxlProjName"

            Start-Autogen20 -CliPath $cliPath -InputInfoPath $inputInfoPath
            Sync-Project20 `
                -AutogenOutputRoot $autogenOutputRootPath `
                -ProjectFolder $oasisProjectFolderPath `
                -IgxlProjName $igxlProjName
            Start-OasisAddin20 -CliPath $addinClientPath
            break
        }
        'Autogen' {
            Start-Autogen20 -CliPath $cliPath -InputInfoPath $inputInfoPath
            break
        }
        'Sync' {
            $igxlProjName = Get-SingleIgxlProj -FolderPath $projectFolder
            Write-Log -Level 'Information' -Message "Target igxlProj: $igxlProjName"

            Sync-Project20 `
                -AutogenOutputRoot $autogenOutputRootPath `
                -ProjectFolder $oasisProjectFolderPath `
                -IgxlProjName $igxlProjName
            break
        }
        'Oasis' {
            Start-OasisAddin20 -CliPath $addinClientPath
            break
        }
    }

    Write-Log -Level 'Information' -Message "[END] '$Mode' mode integration process complete"
}

function Start-Autogen20 {
    param (
        $CliPath,
        $InputInfoPath
    )

    Write-Log -Level 'Information' -Message "[START] Executing AutogenCommandLine..."
    Start-AutogenProcess -CliPath $CliPath -InputInfoPath $InputInfoPath
    Write-Log -Level 'Information' -Message "[DONE] AutogenCommandLine complete!"
}

function Sync-Project20 {
    param (
        $AutogenOutputRoot,
        $ProjectFolder,
        $IgxlProjName
    )

    Write-Log -Level 'Information' -Message "[START] Starting Autogen -> Oasis workspace sync..."
    Sync-AutogenWorkspaceToOasis -AutogenOutputRoot $AutogenOutputRoot -ProjectFolder $ProjectFolder -IgxlProjName $IgxlProjName
    Write-Log -Level 'Information' -Message "[DONE] Autogen -> Oasis workspace sync completed successfully."
}

function Start-OasisAddin20 {
    param (
        $CliPath,
        $IgxlProjName
    )

    Write-Log -Level 'Information' -Message "[START] Executing OasisAddinClient..."
    Start-OasisAddinClient -CliPath $CliPath -Params "IGLinkSyncImport false"
    Write-Log -Level 'Information' -Message "[DONE] OasisAddinClient complete!"
}

# Execute
Invoke-Main
