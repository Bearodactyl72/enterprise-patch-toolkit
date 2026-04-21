# DOTS formatting comment

@{
    RootModule        = 'Test-ConnectionAsJob.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = 'a1b2c3d4-1007-4000-8000-000000000007'
    Author            = 'Skyler Werner'
    Description       = 'Parallel connectivity testing using Test-Connection -AsJob. Returns ComputerName, IPV4Address, and Reachable for each target.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Test-ConnectionAsJob')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
