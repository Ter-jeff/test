@{
    RootModule        = 'AutogenClient.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'd45ea8b4-8f51-4900-9c35-06c3bda4e912'

    Author            = 'Victor Shih'
    CompanyName       = 'Teradyne'
    Description       = 'PowerShell client module for invoking Autogen.'

    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Invoke-Autogen'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    RequiredModules = @(
        'AppCore',
        'AppLogging'
    )

}
