# DOTS formatting comment

@{
    RootModule        = 'Get-RegistryKey.psm1'
    ModuleVersion     = '2.0.0'
    GUID              = 'a1b2c3d4-2002-4000-8000-000000000021'
    Author            = 'Skyler Werner'
    Description       = 'Queries registry uninstall keys on remote machines for matching software entries using runspaces. Returns ComputerName, KeyPresent, DisplayName, and Version.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Get-RegistryKey')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
