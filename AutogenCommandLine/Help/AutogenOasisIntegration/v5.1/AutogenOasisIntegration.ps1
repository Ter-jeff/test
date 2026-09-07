param(
    [ValidateSet('Development', 'Production')]
    [string]$Environment = 'Production',

    [ValidateSet('All', 'Autogen', 'Sync', 'Oasis')]
    [string]$Mode = 'All'
)

$repoRoot = Split-Path -Parent $PSScriptRoot
# Determine environment configuration file paths
$configPath = if ($Environment -eq 'Development') {
    Join-Path $repoRoot 'appsettings.Development.csv'
} else {
    Join-Path $repoRoot 'appsettings.csv'
}

# Import modules
Import-Module AppCore -Force -ErrorAction Stop
Import-Module AppLogging -Force -ErrorAction Stop
Import-Module AutogenClient -Force -ErrorAction Stop
Import-Module AutogenOasisWorkspace -Force -ErrorAction Stop
Import-Module OasisAddinClient -Force -ErrorAction Stop

function Invoke-Main {
    try {
        $repoRoot = Split-Path -Path $PSScriptRoot -Parent

        # Target the central .logs folder at the root level
        $logDirectory = Join-Path -Path $repoRoot -ChildPath ".logs"
        Initialize-Logger -LogDirectory $logDirectory
        Write-Log -Level 'Information' -Message "[BOOT] Running Powershell 5.1 in '$Environment' environment, '$Mode' mode"
        Write-Log -Level 'Information' -Message "Using config file: $configPath"

        Start-Integration
        exit [int][ExitCode]::Success
    }
    catch {
        # Catch block inspection to distinguish expected integration errors from unexpected system crashes
        $errorRecord = $_

        if ($errorRecord.FullyQualifiedErrorId -eq 'IntegrationEngineError' -or
            ($null -ne $errorRecord.Exception -and $errorRecord.Exception.PSObject.Properties['ExitCode'])) {

            Write-ExpectedError $errorRecord

            $exitCodeValue = $errorRecord.Exception.ExitCode
            if ($null -eq $exitCodeValue) { $exitCodeValue = [int][ExitCode]::UnknownError }

            exit [int]$exitCodeValue
        }
        else {
            Write-UnexpectedError $errorRecord
            exit [int][ExitCode]::UnknownError
        }
    }
}

function Start-Integration {
    Write-Log -Level 'Information' -Message "[START] Initializing configuration..."

    # Extract configuration data
    $appSettings = Get-AppSettings -ConfigPath $configPath

    $autogenConfig = $appSettings["autogen"]
    $cliPath       = $autogenConfig["cliPath"]
    $inputInfoPath = $autogenConfig["inputInfoPath"]

    $oasisConfig     = $appSettings["oasis"]
    $addinClientPath = $oasisConfig["addinClientPath"]
    $projectFolder   = $oasisConfig["projectFolder"]

    Write-Log -Level 'Information' -Message "AutogenCommandLine path: $cliPath"
    Write-Log -Level 'Information' -Message "Args path: $inputInfoPath"
    Write-Log -Level 'Information' -Message "OasisAddinClient path: $addinClientPath"
    Write-Log -Level 'Information' -Message "OasisAddinClient projectFolder: $projectFolder"

    # Read inputinfo.csv and get output folder
    $inputInfoMap = Get-InputInfo -CsvPath $inputInfoPath
    $outputFolder = Get-OutputFolder -InputMap $inputInfoMap

    Write-Log -Level 'Information' -Message "Output folder: $outputFolder"
    Write-Log -Level 'Information' -Message "[DONE] Configuration initialized!"

    switch ($Mode) {
        'All' {
            $igxlProjName = Get-SingleIgxlProj -FolderPath $projectFolder
            Write-Log -Level 'Information' -Message "Target igxlProj: $igxlProjName"

            Start-Autogen51 -CliPath $cliPath -InputInfoPath $inputInfoPath
            Sync-Project51 `
                -AutogenOutputRoot "${outputFolder}\IGLink" `
                -ProjectFolder $projectFolder `
                -IgxlProjName $igxlProjName
            Start-OasisAddin51 -CliPath $addinClientPath
            break
        }
        'Autogen' {
            Start-Autogen51 -CliPath $cliPath -InputInfoPath $inputInfoPath
            break
        }
        'Sync' {
            $igxlProjName = Get-SingleIgxlProj -FolderPath $projectFolder
            Write-Log -Level 'Information' -Message "Target igxlProj: $igxlProjName"

            Sync-Project51 `
                -AutogenOutputRoot "${outputFolder}\IGLink" `
                -ProjectFolder $projectFolder `
                -IgxlProjName $igxlProjName
            break
        }
        'Oasis' {
            Start-OasisAddin51 -CliPath $addinClientPath
            break
        }
    }

    Write-Log -Level 'Information' -Message "[END] '$Mode' mode integration process complete!"
}

function Start-Autogen51 {
    param (
        [Parameter(Mandatory)]
        [string]$CliPath,

        [Parameter(Mandatory)]
        [string]$InputInfoPath
    )

    # Run Autogen
    Write-Log -Level 'Information' -Message "[START] Executing AutogenCommandLine..."
    Invoke-Autogen -CliPath $CliPath -InputInfoPath $InputInfoPath
    Write-Log -Level 'Information' -Message "[DONE] AutogenCommandLine complete!"
}

function Sync-Project51 {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$AutogenOutputRoot,

        [Parameter(Mandatory)]
        [string]$ProjectFolder,

        [Parameter(Mandatory)]
        [string]$IgxlProjName
    )

    # Backup project files and sync Autogen output into project path
    Write-Log -Level 'Information' -Message "[START] Starting Autogen -> Oasis workspace sync..."
    Sync-AutogenWorkspaceToOasis `
        -AutogenOutputRoot $AutogenOutputRoot `
        -ProjectFolder $ProjectFolder `
        -IgxlProjName $IgxlProjName
    Write-Log -Level 'Information' -Message "[DONE] Autogen -> Oasis workspace sync completed successfully."
}

function Start-OasisAddin51 {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$CliPath
    )

    # Run OasisAddinClient.exe
    Write-Log -Level 'Information' -Message "[START] Executing OasisAddinClient..."
    Invoke-OasisAddinClient -CliPath $CliPath -Params "IGLinkSyncImport false"
    Write-Log -Level 'Information' -Message "[DONE] OasisAddinClient complete!"
}

function Write-ExpectedError {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    Write-Log -Level 'Error' -Message "[FAIL] Autogen-Oasis integration failed."
    Write-Log -Level 'Error' -Message "$($ErrorRecord.Exception.Message)"
}

function Write-UnexpectedError {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    Write-Log -Level 'Error' -Message "[FAIL] Autogen-Oasis integration: An unexpected exception occurred."
    Write-Log -Level 'Error' -Message "$($ErrorRecord.Exception.Message)"

    if ($ErrorRecord.Exception.InnerException) {
        Write-Log -Level 'Error' -Message "Inner exception: $($ErrorRecord.Exception.InnerException.Message)"
    }

    $errorCustomStackTrace = $ErrorRecord.ScriptStackTrace
    if (-not $errorCustomStackTrace) { $errorCustomStackTrace = $ErrorRecord.StackTrace }

    Write-Log -Level 'Error' -Message "Stack trace:`n$errorCustomStackTrace"
}

# Fire the program entry point
Invoke-Main
