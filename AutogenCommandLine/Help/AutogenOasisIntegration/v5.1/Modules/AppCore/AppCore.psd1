@{
    RootModule        = 'AppCore.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '3d8a2a2a-d8c1-4b5a-9e25-cc0db8a4521f'

    Author            = 'Victor Shih'
    CompanyName       = 'Teradyne'
    Description       = 'Core domain types and configuration loader for Autogen/Oasis orchestration.'

    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Get-AppSettings',
        'Get-InputInfo',
        'Get-OutputFolder',
        'New-IntegrationError',
        'Get-SingleIgxlProj'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
