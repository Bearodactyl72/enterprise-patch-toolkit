# DOTS formatting comment

@{
    RootModule        = 'Get-DeepClone.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-1005-4000-8000-000000000005'
    Author            = 'Skyler Werner'
    Description       = 'Creates independent deep copies of hashtables to prevent reference leakage in concurrent operations.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Get-DeepClone')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
