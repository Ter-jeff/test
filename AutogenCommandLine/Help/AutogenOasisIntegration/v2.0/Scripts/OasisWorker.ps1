function Start-OasisAddinClient {
    param (
        $CliPath,
        $Params
    )

    if (-not $CliPath -or -not $Params) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message "Missing mandatory parameters CliPath or Params inside Start-OasisProcess." -ExitCodeValue $errCode)
    }

    if (-not (Test-Path $CliPath)) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message ("Oasis executable missing at path: " + $CliPath) -ExitCodeValue $errCode)
    }

    $folderPath = Split-Path -Path $CliPath -Parent
    Push-Location $folderPath

    try {
        $process = Start-Process `
            -FilePath $CliPath `
            -ArgumentList $Params `
            -WorkingDirectory $folderPath `
            -NoNewWindow `
            -Wait `
            -PassThru

        $exitCode = $process.ExitCode
    }
    catch {
        $errCode = $Global:ExitCodeMap.UnknownError
        $failMessage = "Fatal error executing external Oasis framework: " + $_.ToString()
        throw (New-IntegrationError -Message $failMessage -ExitCodeValue $errCode)
    }
    finally {
        if ($null -ne $process) { $process.Dispose() }
        Pop-Location
    }

    if ($exitCode -ne 0) {
        $errCode = $Global:ExitCodeMap.UnknownError
        $failMessage = "OasisAddInClient failed with exit code " + $exitCode + "."
        throw (New-IntegrationError -Message $failMessage -ExitCodeValue $errCode)
    }
}
