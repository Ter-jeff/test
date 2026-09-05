@{
    RootModule        = 'AutogenOasisWorkspace.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '1c7f5d1f-6a6a-4c17-991d-9ddfd58b1c78'

    Author            = 'Victor Shih'
    CompanyName       = 'Teradyne'
    Description       = 'Synchronize Autogen output into Oasis workspace with validation and backup.'

    PowerShellVersion = '5.1'

    RequiredModules = @(
        'AppCore'
        'AppLogging'
    )

    FunctionsToExport = @(
        'Sync-AutogenWorkspaceToOasis'
    )
}
