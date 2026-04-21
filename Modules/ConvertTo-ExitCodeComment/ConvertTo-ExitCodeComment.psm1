# DOTS formatting comment

<#
    .SYNOPSIS
        Converts a process exit code to a human-readable comment string.
    .DESCRIPTION
        Converts a process exit code to a human-readable comment string.
        Covers common Windows Installer, Windows Update, DISM, and PsExec
        exit codes encountered during software patching operations.
        Written by Skyler Werner
        Date: 2026/03/23
        Version 1.0.0
#>

function ConvertTo-ExitCodeComment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [int]$ExitCode
    )

    switch ($ExitCode) {

        # --- Standard Windows / MSI exit codes ---
        0       { "Completed Successfully" }
        1       { "Incorrect Function" }
        2       { "File cannot be found" }
        3       { "Cannot find the specified path" }
        5       { "Access denied" }
        8       { "Not enough memory resources are available to process this command" }
        13      { "The data is invalid" }
        23      { "The component store has been corrupted / Generic Error - Check logs" }
        38      { "Windows is unable to load the device driver because a previous version is still in memory, resulting in conflicts" }
        53      { "The network path was not found" }
        59      { "Unexpected network error" }
        87      { "The parameter is incorrect" }
        112     { "There is not enough space on the disk." }
        184     { "A necessary file is locked by another process" }
        233     { "No process is on the other end of the pipe" }
        267     { "Directory name is invalid" }
        1060    { "The specified service does not exist as an installed service" }
        1392    { "A file or files are corrupt" }
        1450    { "Insufficient system resources exist to complete the requested service" }
        1602    { "User canceled installation" }
        1603    { "Fatal error during installation" }
        1605    { "This action is only valid for products that are currently installed" }
        1612    { "The installation source for this product is not available" }
        1618    { "Another installation is in progress" }
        1619    { "The installation package could not be opened" }
        1620    { "This installation package could not be opened. Contact the application vendor to verify that this is a valid Windows Installer package." }
        1635    { "This update package could not be opened" }
        1636    { "Update package cannot be opened" }
        1641    { "Successful - Restart required" }
        1642    { "Upgrade cannot be installed - Missing software" }
        1648    { "No valid sequence could be found for the set of patches" }
        1726    { "The remote procedure call failed" }
        3010    { "Successful - Restart required" }
        14001   { "The application failed to start" }
        14098   { "The component store has been corrupted" }
        2359302 { "The patch has already been installed" }

        # --- Negative HRESULT / Windows Update / DISM codes ---
        -1073741510 { "Cmd.exe window was closed" }                                                             # 0xC000013A
        -2067919934 { "SQL server related error" }                                                              # 0x84BE0BC2
        -2145116147 { "The update handler did not install the update because it needs to be downloaded again" } # 0x8024200D
        -2146959355 { ".NET framework installation failure" }                                                   # 0x80080005
        -2145124329 { "Patch installation failure" }                                                            # 0x80240017
        -2145124322 { "Generic error" }                                                                         # 0x8024001E
        -2145124330 { "Another install is ongoing or reboot is pending" }                                       # 0x80240016
        -2146498167 { "The device is missing important security and quality fixes" }                            # 0x800F0989
        -2146498172 { "The matching component directory exists but binary is missing" }                         # 0x800F0984
        -2146498299 { "DISM Package Manager processed the command line but failed" }                            # 0x800F0905
        -2146498304 { "An unknown error occurred" }                                                             # 0x800F0900
        -2146498511 { "Corruption in the windows component store" }                                             # 0x800F0831
        -2147417839 { "OLE received a packet with an invalid header (RPC Error)" }                              # 0x80010111
        -2147467259 { "A file that the Windows Product Activation (WPA) requires is damaged or missing" }       # 0x80004005
        -2147956498 { "The component store has been corrupted" }

        default { $null }
    }
}
