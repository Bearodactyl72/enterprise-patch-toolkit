# DOTS formatting comment

@{
    RootModule        = 'TargetMachines.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-2004-4000-8000-000000000040'
    Author            = 'Skyler Werner'
    Description       = 'Manages the global $TargetMachines list -- import from file or enter interactively.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Import-TargetMachines', 'Enter-TargetMachines')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @('Load-TargetMachines')
}
