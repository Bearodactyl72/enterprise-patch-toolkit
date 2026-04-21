# DOTS formatting comment

@{
    RootModule        = 'Get-DriverInfo.psm1'
    ModuleVersion     = '2.0.0'
    GUID              = 'a1b2c3d4-2001-4000-8000-000000000030'
    Author            = 'Skyler Werner, Alec Barrett'
    Description       = 'Retrieves driver information from remote machines using runspaces. Supports .inf drivers (via Get-WindowsDriver) and loaded drivers (via driverquery). Uses Invoke-RunspacePool for concurrent execution.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Get-DriverInfo')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
