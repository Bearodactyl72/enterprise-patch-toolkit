# DOTS formatting comment

@{
    RootModule        = 'Invoke-RunspacePool.psm1'
    ModuleVersion     = '2.0.1'
    GUID              = 'a1b2c3d4-1001-4000-8000-000000000001'
    Author            = 'Skyler Werner'
    Description       = 'Concurrent execution engine using RunspacePool. Replaces legacy *AsJob patterns with timeout handling and progress monitoring.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Invoke-RunspacePool', 'Stop-RunspaceAsync')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
