# DOTS formatting comment

@{
    RootModule        = 'Format-ComputerList.psm1'
    ModuleVersion     = '1.0.1'
    GUID              = 'a1b2c3d4-1002-4000-8000-000000000002'
    Author            = 'Skyler Werner'
    Description       = 'Sanitizes computer name lists by removing domain prefixes/suffixes and duplicates based on the active network profile from RSL-Environment.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Format-ComputerList')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
