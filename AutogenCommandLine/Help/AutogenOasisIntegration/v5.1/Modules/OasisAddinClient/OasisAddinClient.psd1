@{
    RootModule        = 'OasisAddinClient.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b6a6d8bb-93dd-4e39-8b9f-0fa0832a5c4f'

    Author            = 'Victor Shih'
    CompanyName       = 'Teradyne'
    Description       = 'PowerShell client module for invoking OasisAddinClient.'

    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Invoke-OasisAddinClient'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    RequiredModules = @(
        'AppCore',
        'AppLogging'
    )
}
