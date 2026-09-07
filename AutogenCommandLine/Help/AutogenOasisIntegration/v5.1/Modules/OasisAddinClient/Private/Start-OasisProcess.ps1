function Start-OasisProcess {
    param (
        [Parameter(Mandatory)]
        [string]$CliPath,

        [Parameter(Mandatory)]
        [string]$Params
    )

    # Clean parameter array breakdown for the arguments string
    $argumentList = $Params -split '\s+'

    # 1. Setup process start info
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $CliPath
    $psi.Arguments = $argumentList
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
            Write-Log -Level Information -Message "$($EventArgs.Data)" -Scope "Oasis"
        }
    }

    $errorAction = {
        if ($EventArgs.Data) {
            Write-Log -Level Error -Message "$($EventArgs.Data)" -Scope "Oasis"
        }
    }

    # 3. Register standard streams into the active PowerShell runspace engine
    $outEvent = Register-ObjectEvent -InputObject $process -EventName "OutputDataReceived" -Action $outputAction
    $errEvent = Register-ObjectEvent -InputObject $process -EventName "ErrorDataReceived" -Action $errorAction

    try {
        [void]$process.Start()
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()

        while (-not $process.HasExited) {
            Start-Sleep -Milliseconds 10
        }

        $process.WaitForExit()
        $exitCode = $process.ExitCode
    }
    finally {
        # Tear down events and discard the handle to avoid memory leaks
        Unregister-Event -SourceIdentifier $outEvent.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $errEvent.Name -ErrorAction SilentlyContinue
        $process.Dispose()
    }

    if ($exitCode -ne 0) {
        throw (New-IntegrationError `
            -Message "OasisAddInClient failed with exit code $($exitCode)." `
            -ExitCode ([ExitCode]::OasisAddinClientFailed))
    }
}
