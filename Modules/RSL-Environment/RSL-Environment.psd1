# DOTS formatting comment

@{
    RootModule        = 'RSL-Environment.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-1002-4000-8000-000000000010'
    Author            = 'Skyler Werner'
    Description       = 'Loads environment-specific configuration and resolves the active network profile for the Remediation Script Library.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Import-RSLEnvironment', 'Get-RSLActiveNetwork')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
