@{
    # ---- Module identity ----
    RootModule        = 'AppLogging.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '6e2b7a2c-20af-4a56-9a4b-dc9b8b69d6fe'

    Author            = 'Victor Shih'
    CompanyName       = 'Teradyne'
    Description       = 'Centralized application logger with file and console output.'

    # ---- Compatibility ----
    PowerShellVersion = '5.1'
    CompatiblePSEditions = @('Desktop','Core')

    # ---- Exports ----
    FunctionsToExport = @(
        'Initialize-Logger',
        'Write-Log'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
