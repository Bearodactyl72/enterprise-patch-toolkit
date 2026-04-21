# DOTS formatting comment

@{
    RootModule        = 'Merge-MainSwitch.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-2003-4000-8000-000000000030'
    Author            = 'Skyler Werner'
    Description       = 'Content-aware merge tool for Main-Switch.ps1. Compares local and central copies case-by-case, auto-merges non-conflicting changes, and prompts for conflict resolution. Supports Receive-MainSwitch (pull), Submit-MainSwitch (push), and Compare-MainSwitch (read-only diff).'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Receive-MainSwitch', 'Submit-MainSwitch', 'Compare-MainSwitch')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @('pull-mainswitch', 'push-mainswitch')
}
