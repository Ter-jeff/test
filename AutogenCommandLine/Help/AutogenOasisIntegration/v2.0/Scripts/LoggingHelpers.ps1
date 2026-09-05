# Script-scoped tracking variable for our log file path destination
$Script:LogFilePath = $null

function Initialize-Logger {
    param(
        $LogDirectory
    )

    # Create the log folder safely if it doesn't exist
    if (-not (Test-Path $LogDirectory)) {
        # New-Item works seamlessly back to 2.0 for directories
        $null = New-Item -Path $LogDirectory -ItemType "Directory" -Force
    }

    $timestamp = [DateTime]::Now.ToString("yyyyMMddHHmmss")
    $logFileName = "log-20-" + $timestamp + ".log"

    $Script:LogFilePath = Join-Path $LogDirectory $logFileName
}

function Write-Log {
    param(
        $Level = 'Information',
        $Message
    )

    if ($null -eq $Script:LogFilePath) {
        $repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
        Initialize-Logger -LogDirectory (Join-Path $repoRoot ".logs")
    }

    $timeString = [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[{0}][{1}] {2}" -f $timeString, $Level, $Message

    if ($Level -eq 'Error') {
        Write-Host $logLine -ForegroundColor Red
    }
    else {
        Write-Host $logLine -ForegroundColor Cyan
    }

    try {
        $logLine | Out-File -FilePath $Script:LogFilePath -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "[WARNING] Unable to append string line write lock to log file path destination." -ForegroundColor Yellow
    }
}

function Write-ExpectedError {
    param($ErrorRecord)

    $msg = "[FAIL] Integration Core Exception: " + $ErrorRecord.Exception.Message
    Write-Log -Level 'Error' -Message $msg
}

function Write-UnexpectedError {
    param($ErrorRecord)

    $msg = "[CRITICAL] An unexpected environment crash occurred: " + $ErrorRecord.ToString()
    if ($null -ne $ErrorRecord.ScriptStackTrace) {
        $msg += " | Stack Trace: " + $ErrorRecord.ScriptStackTrace
    }
    Write-Log -Level 'Error' -Message $msg
}
