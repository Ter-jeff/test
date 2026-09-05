function Start-AutogenProcess {
    param (
        $CliPath,
        $InputInfoPath
    )

    if (-not $CliPath -or -not $InputInfoPath) {
        $errCode = $Global:ExitCodeMap.UnknownError
        throw (New-IntegrationError -Message "Missing mandatory parameters CliPath or InputInfoPath inside Start-AutogenProcess." -ExitCodeValue $errCode)
    }

    $autogenArgs = "";
    $raw = Import-Csv $InputInfoPath
    foreach ($row in $raw) {
        $autogenArgs = $autogenArgs + $row.Item + " " + $row.Value + " "
    }

    $folderPath = Split-Path -Path $CliPath -Parent
    Push-Location $folderPath

    try {
        $process = Start-Process -FilePath $CliPath `
            -ArgumentList $autogenArgs `
            -WorkingDirectory $folderPath `
            -NoNewWindow `
            -Wait `
            -PassThru

        $exitCode = $process.ExitCode
    }
    catch {
        $errCode = $Global:ExitCodeMap.UnknownError
        $failMessage = "Fatal error executing external process framework: " + $_.ToString()
        throw (New-IntegrationError -Message $failMessage -ExitCodeValue $errCode)
    }
    finally {
        if ($null -ne $process) { $process.Dispose() }
        Pop-Location
    }

    if ($exitCode -ne 0) {
        $errCode = $Global:ExitCodeMap.UnknownError
        $failMessage = "Autogen CLI failed with exit code " + $exitCode + ". Please see Autogen logs at " + $folderPath + "\Teradyne"
        throw (New-IntegrationError -Message $failMessage -ExitCodeValue $errCode)
    }
}
