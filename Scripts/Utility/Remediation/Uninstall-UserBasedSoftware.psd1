# DOTS formatting comment

<#
    --- Field Reference ---

    ProcessName   : Wildcard pattern for taskkill (e.g., "Grammarly.*", "Zoom*").
                    Supports wildcards. Used to force-stop the application before uninstall.

    DiscoveryPath : Path relative to the user profile folder (C:\Users\<user>\...).
                    Points to the directory containing the discovery executable.

    DiscoveryExe  : The main executable used to detect whether the software is installed.
                    The function checks for this file under each user profile to find
                    installations and reads its VersionInfo for version reporting.

    UninstallPath : Path relative to the user profile folder (C:\Users\<user>\...).
                    Points to the directory containing the uninstaller executable.
                    Can be the same as DiscoveryPath or a different location.

    UninstallExe  : The uninstaller executable found at UninstallPath.
                    This is what the scheduled task actually runs.

    UninstallArgs : Silent/quiet arguments for the uninstaller (/S, /silent, /quiet, etc.).
                    Passed directly to UninstallExe on the command line.
#>

@{
    Grammarly = @{
        ProcessName   = "Grammarly.*"
        DiscoveryPath = "AppData\Local\Grammarly\DesktopIntegrations"
        DiscoveryExe  = "Grammarly.Desktop.exe"
        UninstallPath = "AppData\Local\Grammarly\DesktopIntegrations"
        UninstallExe  = "Uninstall.exe"
        UninstallArgs = "/S"
    }

    Zoom = @{
        ProcessName   = "Zoom*"
        DiscoveryPath = "AppData\Roaming\Zoom\bin"
        DiscoveryExe  = "Zoom.exe"
        UninstallPath = "AppData\Roaming\Zoom\bin"
        UninstallExe  = "Uninstall.exe"
        UninstallArgs = "/S"
    }

    Spotify = @{
        ProcessName   = "Spotify*"
        DiscoveryPath = "AppData\Roaming\Spotify"
        DiscoveryExe  = "Spotify.exe"
        UninstallPath = "AppData\Roaming\Spotify"
        UninstallExe  = "Spotify.exe"
        UninstallArgs = "/uninstall /S"
    }
}
