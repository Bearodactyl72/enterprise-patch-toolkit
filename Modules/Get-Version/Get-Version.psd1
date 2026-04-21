# DOTS formatting comment

@{
    RootModule        = 'Get-Version.psm1'
    ModuleVersion     = '2.0.0'
    GUID              = 'a1b2c3d4-2001-4000-8000-000000000020'
    Author            = 'Skyler Werner'
    Description       = 'Retrieves file version info from remote machines using runspaces. Supports user-profile paths, hidden files, and ProductVersion mode. Returns ComputerName, Version, Installed, and ResolvedPath.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Get-Version')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
