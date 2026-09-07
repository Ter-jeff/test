function Start-AutogenProcess {
    param (
        [Parameter(Mandatory)]
        [string]$CliPath,

        [Parameter(Mandatory)]
        [string]$InputInfoPath
    )

    $autogenArgs = "";
    $raw = Import-Csv $InputInfoPath
    foreach ($row in $raw) {
        $autogenArgs = $autogenArgs + $row.Item + " " + $row.Value + " "
    }

    $folderPath = Split-Path -Path $CliPath -Parent
    Push-Location $folderPath

    # Setup process start info
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $CliPath
    $psi.Arguments = [string]::Join(" ", $autogenArgs)
    $psi.WorkingDirectory = $folderPath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    # Define output actions
    $outputAction = {
        if ($EventArgs.Data) {
            Write-Log -Level Information -Message "$($EventArgs.Data)" -Scope "Autogen"
        }
    }

    $errorAction = {
        if ($EventArgs.Data) {
            Write-Log -Level Error -Message "$($EventArgs.Data)" -Scope "Autogen"
        }
    }

    # Register events into the active runspace
    $outEvent = Register-ObjectEvent -InputObject $process -EventName "OutputDataReceived" -Action $outputAction
    $errEvent = Register-ObjectEvent -InputObject $process -EventName "ErrorDataReceived" -Action $errorAction

    try {
        # Start the process and begin async reading
        [void]$process.Start()
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()

        # Wait loop keeping the PowerShell event engine processing safely
        while (-not $process.HasExited) {
            Start-Sleep -Milliseconds 10
        }

        $process.WaitForExit()
        $exitCode = $process.ExitCode
    }
    finally {
        # Clean up resource allocations
        Unregister-Event -SourceIdentifier $outEvent.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $errEvent.Name -ErrorAction SilentlyContinue
        $process.Dispose()

        Pop-Location
    }

    if ($exitCode -ne 0) {
        throw (New-IntegrationError `
            -Message "Autogen CLI failed with exit code $($exitCode). Please see Autogen logs at ${folderPath}\Teradyne" `
            -ExitCode ([ExitCode]::AutogenFailed))
    }
}
