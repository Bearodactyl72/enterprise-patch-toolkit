# DOTS formatting comment

function Start-UserProfileCleanup {
    <#
        .SYNOPSIS
            Removes stale, empty, and temporary user profiles from remote machines.
        .DESCRIPTION
            Scans C:\Users on each target machine and removes profiles that are old,
            empty, or temporary (Windows creates .DOMAIN-suffixed temp profile
            folders when a domain profile fails to load). Also cleans orphaned
            ProfileList registry
            entries, stale .bak ProfileList keys, ProfileGuid orphans, and per-user
            Group Policy cache entries.

            Protected profiles are never touched:
            - Active or disconnected sessions (query user)
            - Loaded profiles (CIM Loaded property, re-checked at deletion time)
            - CIM Special profiles (system-critical)
            - Profiles with mandatory or in-creation registry State flags
            - Profiles whose accounts run Windows services
            - Profiles whose accounts run scheduled tasks

            Use -WhatIf to preview what would be removed without making changes.
        .PARAMETER ComputerName
            One or more computer names to clean up. Accepts pipeline input.
        .PARAMETER MinFreeGB
            Minimum free disk space in GB. Machines above this threshold are skipped.
            Default: 10
        .PARAMETER UserMaxAgeDays
            Profiles with no activity within this many days are considered old.
            Default: 500
        .PARAMETER TempMaxAgeDays
            Temp profiles (.DOMAIN suffix) with no activity within this many days
            are removed. Default: 90
        .PARAMETER ThrottleLimit
            Maximum concurrent machines to process. Default: 25
        .PARAMETER TimeoutMinutes
            Minutes before a machine's cleanup task is auto-stopped. Default: 30
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines2.txt"
            Start-UserProfileCleanup -ComputerName $list
        .EXAMPLE
            Start-UserProfileCleanup -ComputerName $list -WhatIf
        .EXAMPLE
            Start-UserProfileCleanup -ComputerName "PC01","PC02" -MinFreeGB 5 -UserMaxAgeDays 365
        .NOTES
            Written by Skyler Werner
            Date: 2026/03/23
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateRange(1, 100)]
        [int]
        $MinFreeGB = 10,

        [Parameter()]
        [ValidateRange(1, 3650)]
        [int]
        $UserMaxAgeDays = 500,

        [Parameter()]
        [ValidateRange(1, 3650)]
        [int]
        $TempMaxAgeDays = 90,

        [Parameter()]
        [ValidateRange(1, 100)]
        [int]
        $ThrottleLimit = 25,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 30,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        $collectedNames = @()
    }

    process {
        foreach ($name in $ComputerName) {
            if ($name.Length -gt 0) {
                $collectedNames += $name
            }
        }
    }

    end {

    $list = Format-ComputerList $collectedNames -ToUpper

    # Capture WhatIf preference as a simple boolean to pass into remote sessions
    $dryRun = $WhatIfPreference


    $scriptblock = {
        $computer = $args[0]
        $size     = $args[1]
        $userAge  = $args[2]
        $tempAge  = $args[3]
        $dryRun   = $args[4]

        $remoteBlock = {
            param($size, $userAge, $tempAge, $dryRun)

            $path = "C:\Users"
            $cutOffDate = (Get-Date).AddDays(-$userAge)
            $tempCutDate = (Get-Date).AddDays(-$tempAge)
            $regProfiles = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
            $regProfileGuid = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileGuid"
            $regGPState = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State"

            # --- Check disk space ---
            $disk = Get-PSDrive ($env:SystemDrive).Replace(":", "")
            $diskFreeGB = [math]::Round($($disk.Free / 1gb), 2)

            if ($diskFreeGB -gt $size) {
                return [PSCustomObject]@{
                    OldRemoved    = $null
                    EmptyRemoved  = $null
                    TempRemoved   = $null
                    OrphanRemoved = $null
                    BakRemoved    = $null
                    GuidRemoved   = $null
                    GPRemoved     = $null
                    OldNames      = @()
                    EmptyNames    = @()
                    TempNames     = @()
                    OrphanNames   = @()
                    BakNames      = @()
                    GuidNames     = @()
                    GPNames       = @()
                    SkippedUsers  = @()
                    SpaceFreed    = $null
                    TotalFree     = $diskFreeGB
                    Comment       = "Cleanup not initiated"
                }
            }

            # --- Exclusion patterns for usernames ---
            # Skip the current admin, built-ins, and common service-account
            # name shapes. Tune these for your environment's naming conventions.
            $excludePatterns = @(
                "admin"
                "svc"
                "helpdesk"
                "Public"
                "Default"
                "All Users"
                $env:USERNAME
            )

            # Known subfolders that indicate profile activity. Org-branded
            # OneDrive folders (OneDrive - <Tenant>) are matched by glob
            # below; no need to hardcode tenant names here.
            $activityFolders = @(
                "Desktop"
                "Documents"
                "Downloads"
                "OneDrive"
                "Pictures"
                "Videos"
            )

            # Pattern for temp profile folders Windows creates when a domain
            # profile fails to load: <user>.<NetBIOS-domain>
            $tempProfilePattern = if ($env:USERDOMAIN) { "\.$env:USERDOMAIN$" } else { $null }


            # --- Get CIM profiles once; build lookups ---
            $cimProfiles = @(Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue)
            $cimByPath = @{}
            foreach ($cim in $cimProfiles) {
                if ($cim.LocalPath) {
                    $cimByPath[$cim.LocalPath] = $cim
                }
            }

            # Build set of loaded profile paths -- never touch these
            $loadedPaths = @{}
            foreach ($cim in $cimProfiles) {
                if ($cim.Loaded -and $cim.LocalPath) {
                    $loadedPaths[$cim.LocalPath] = $true
                }
            }

            # Build set of Special (system-critical) profile paths
            $specialPaths = @{}
            foreach ($cim in $cimProfiles) {
                if ($cim.Special -and $cim.LocalPath) {
                    $specialPaths[$cim.LocalPath] = $true
                }
            }

            # Build set of profiles with protected registry State flags.
            # State 0x0001 = mandatory profile (shared template, never delete)
            # State 0x0002 = new profile (first logon in progress)
            $protectedStatePaths = @{}
            $regKeys = @(Get-ChildItem $regProfiles -ErrorAction SilentlyContinue)
            foreach ($key in $regKeys) {
                $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                if (-not $props.ProfileImagePath) { continue }
                $state = $props.State
                if ($null -eq $state) { continue }
                # Bitwise check for mandatory (0x1) or new/creating (0x2)
                if (($state -band 0x0001) -or ($state -band 0x0002)) {
                    $protectedStatePaths[$props.ProfileImagePath] = $true
                }
            }


            # --- Detect active/disconnected sessions via query user ---
            $activeSessionUsers = @{}
            try {
                $quserOutput = @(query user 2>&1)
                foreach ($line in $quserOutput) {
                    if ($line -match 'USERNAME') { continue }
                    if ($line -match '^\s*>?(\S+)') {
                        $activeSessionUsers[$Matches[1]] = $true
                    }
                }
            }
            catch {
                # query user may not be available -- continue without it
            }


            # --- Detect service and scheduled task accounts ---
            # These profiles must not be deleted even if they appear stale,
            # because the service or task depends on the profile existing.
            $serviceAccountUsers = @{}

            # Services: extract username portion from StartName (DOMAIN\user or .\user)
            try {
                $services = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
                    Where-Object { $_.StartName } |
                    Where-Object { $_.StartName -notmatch "^(LocalSystem|NT AUTHORITY|NT SERVICE)" })
                foreach ($svc in $services) {
                    $svcUser = $svc.StartName
                    # Strip domain prefix (DOMAIN\ or .\)
                    if ($svcUser -match '\\(.+)$') {
                        $svcUser = $Matches[1]
                    }
                    # Strip @domain suffix (user@domain.com)
                    if ($svcUser -match '^([^@]+)@') {
                        $svcUser = $Matches[1]
                    }
                    if ($svcUser) {
                        $serviceAccountUsers[$svcUser] = "service: $($svc.Name)"
                    }
                }
            }
            catch {}

            # Scheduled tasks: extract username from Principal
            try {
                $tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue |
                    Where-Object { $_.Principal.UserId } |
                    Where-Object { $_.State -ne 'Disabled' } |
                    Where-Object { $_.Principal.UserId -notmatch "^(SYSTEM|LOCAL SERVICE|NETWORK SERVICE|S-1-5-)" } |
                    Where-Object { $_.Principal.UserId -notmatch "^(INTERACTIVE|Everyone|Users|Administrators)$" })
                foreach ($task in $tasks) {
                    $taskUser = $task.Principal.UserId
                    if ($taskUser -match '\\(.+)$') {
                        $taskUser = $Matches[1]
                    }
                    if ($taskUser -match '^([^@]+)@') {
                        $taskUser = $Matches[1]
                    }
                    if ($taskUser) {
                        $serviceAccountUsers[$taskUser] = "task: $($task.TaskName)"
                    }
                }
            }
            catch {}


            # Helper: get filtered user directories
            function Get-UserDirectories {
                $dirs = @(Get-ChildItem $path -Directory -Force -ErrorAction SilentlyContinue)
                foreach ($pattern in $excludePatterns) {
                    $dirs = @($dirs | Where-Object { $_.Name -notmatch $pattern })
                }
                # Exclude users with active or disconnected sessions
                $dirs = @($dirs | Where-Object { -not $activeSessionUsers.ContainsKey($_.Name) })
                # Exclude profiles whose CIM entry shows as Loaded
                $dirs = @($dirs | Where-Object { -not $loadedPaths.ContainsKey($_.FullName) })
                # Exclude CIM Special (system-critical) profiles
                $dirs = @($dirs | Where-Object { -not $specialPaths.ContainsKey($_.FullName) })
                # Exclude profiles with mandatory or in-creation State flags
                $dirs = @($dirs | Where-Object { -not $protectedStatePaths.ContainsKey($_.FullName) })
                # Exclude service and scheduled task accounts
                $dirs = @($dirs | Where-Object { -not $serviceAccountUsers.ContainsKey($_.Name) })
                return $dirs
            }


            # Helper: remove a user profile cleanly via CIM, with manual fallback.
            # In dry-run mode, validates that removal WOULD proceed and returns $true
            # without making changes.
            function Remove-UserProfile {
                param([System.IO.DirectoryInfo]$UserDir)

                $localPath = $UserDir.FullName

                # Safety: never delete a currently loaded profile (cached check)
                if ($loadedPaths.ContainsKey($localPath)) { return $false }

                # Safety: re-query CIM right before deletion to catch users who
                # logged in after the initial snapshot was taken
                $escapedPath = $localPath -replace "\\", "\\\\"
                $freshCim = Get-CimInstance Win32_UserProfile -Filter "LocalPath = '$escapedPath'" -ErrorAction SilentlyContinue
                if ($freshCim.Loaded) { return $false }
                if ($freshCim.Special) { return $false }

                # In dry-run mode, stop here -- the profile WOULD be removed
                if ($dryRun) { return $true }

                $removed = $false

                # Primary: CIM deletion (removes folder + registry keys atomically,
                # including ProfileGuid mapping)
                if ($freshCim) {
                    try {
                        Remove-CimInstance -InputObject $freshCim -ErrorAction Stop
                        $removed = $true
                    }
                    catch {
                        # CIM failed -- fall through to manual cleanup
                    }
                }

                # Fallback: manual removal if CIM did not handle it
                if (-not $removed) {
                    Remove-Item $localPath -Force -Recurse -ErrorAction SilentlyContinue

                    # Remove matching ProfileList registry entry (exact path match)
                    $plKeys = @(Get-ChildItem $regProfiles -ErrorAction SilentlyContinue)
                    foreach ($key in $plKeys) {
                        $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                        if ($props.ProfileImagePath -eq $localPath) {
                            # Capture the SID from the key name for related cleanup
                            $sid = Split-Path $key.PSPath -Leaf
                            Remove-Item $key.PSPath -Force -Recurse -ErrorAction SilentlyContinue

                            # Clean ProfileGuid entry that references this SID
                            if (Test-Path $regProfileGuid) {
                                $guidKeys = @(Get-ChildItem $regProfileGuid -ErrorAction SilentlyContinue)
                                foreach ($gk in $guidKeys) {
                                    $gkProps = Get-ItemProperty $gk.PSPath -ErrorAction SilentlyContinue
                                    if ($gkProps.SidString -eq $sid) {
                                        Remove-Item $gk.PSPath -Force -Recurse -ErrorAction SilentlyContinue
                                    }
                                }
                            }
                        }
                    }
                }

                return $true
            }


            # Helper: resolve the SID from a ProfileList registry key to a username.
            # Returns the key leaf name (SID) if resolution fails.
            function Resolve-SidToName {
                param([string]$Sid)
                try {
                    $objSid = New-Object System.Security.Principal.SecurityIdentifier($Sid)
                    $objUser = $objSid.Translate([System.Security.Principal.NTAccount])
                    return $objUser.Value
                }
                catch {
                    return $Sid
                }
            }


            # --- Phase 1: Delete old user profiles ---
            $oldRemoved = @()
            $userDirs = Get-UserDirectories

            foreach ($user in $userDirs) {
                # Skip temp profiles -- handled in Phase 3
                if ($tempProfilePattern -and $user.Name -match $tempProfilePattern) { continue }

                # Check NTUSER.DAT -- updated on every logoff, most reliable single
                # indicator of when the profile was last actively used
                $ntUserDat = Join-Path $user.FullName "NTUSER.DAT"
                if (Test-Path $ntUserDat) {
                    $ntUserTime = (Get-Item $ntUserDat -Force -ErrorAction SilentlyContinue).LastWriteTime
                    if ($ntUserTime -gt $cutOffDate) { continue }
                }

                # Check CIM LastUseTime -- maintained by Windows independently of
                # file system timestamps
                $cimEntry = $cimByPath[$user.FullName]
                if ($cimEntry -and $cimEntry.LastUseTime) {
                    if ($cimEntry.LastUseTime -gt $cutOffDate) { continue }
                }

                $recentFolders = @(
                    Get-ChildItem $user.FullName -Force -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -in $activityFolders } |
                    Where-Object { $_.LastWriteTime -gt $cutOffDate }
                )

                # Org-branded OneDrive folder (OneDrive - <Tenant>); first
                # match wins. Works for any tenant display name.
                $oneDrivePath = @(Get-ChildItem $user.FullName -Directory -Force -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -like 'OneDrive - *' } |
                    Select-Object -First 1 -ExpandProperty FullName)
                $recentOneDrive = @()
                if ($oneDrivePath) {
                    $recentOneDrive = @(
                        Get-ChildItem $oneDrivePath -Force -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -in $activityFolders } |
                        Where-Object { $_.LastWriteTime -gt $cutOffDate }
                    )
                }

                if ($recentFolders.Count -lt 1 -and $recentOneDrive.Count -lt 1) {
                    $didRemove = Remove-UserProfile -UserDir $user
                    if ($didRemove) {
                        $oldRemoved += $user.FullName
                    }
                }
            }


            # --- Phase 2: Delete empty user profiles ---
            # Scans the ENTIRE profile for user-created content, not just activity
            # folders. Excludes default Windows skeleton files so a fresh profile
            # with no real data is correctly identified as empty.
            $emptyRemoved = @()
            $userDirs = Get-UserDirectories

            # Paths that exist in every default profile -- not evidence of real usage
            $defaultPathPatterns = @(
                "\\NTUSER\."
                "\\ntuser\."
                "\\AppData\\Local\\Microsoft\\Windows\\"
                "\\AppData\\Local\\Microsoft\\WindowsApps\\"
                "\\AppData\\Local\\Temp\\"
                "\\AppData\\Roaming\\Microsoft\\Windows\\"
                "\\AppData\\Roaming\\Microsoft\\Internet Explorer\\"
                "\\AppData\\Local\\Microsoft\\Internet Explorer\\"
            )

            foreach ($user in $userDirs) {
                if ($tempProfilePattern -and $user.Name -match $tempProfilePattern) { continue }

                # Skip profiles already flagged as old in Phase 1
                if ($user.FullName -in $oldRemoved) { continue }

                # Search entire profile for any non-default, non-shortcut file.
                # Stop at the first match for performance.
                $hasContent = $false
                $allFiles = @(Get-ChildItem $user.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                    Where-Object { -not $_.PSIsContainer })

                foreach ($file in $allFiles) {
                    if ($file.Extension -eq ".lnk") { continue }
                    if ($file.Name -match "^desktop\.ini$") { continue }

                    $isDefault = $false
                    foreach ($pattern in $defaultPathPatterns) {
                        if ($file.FullName -match $pattern) {
                            $isDefault = $true
                            break
                        }
                    }
                    if ($isDefault) { continue }

                    $hasContent = $true
                    break
                }

                if (-not $hasContent) {
                    $didRemove = Remove-UserProfile -UserDir $user
                    if ($didRemove) {
                        $emptyRemoved += $user.FullName
                    }
                }
            }


            # --- Phase 3: Delete old temp profiles (.<NetBIOS-domain>) ---
            $tempRemoved = @()
            $userDirs = Get-UserDirectories
            if ($tempProfilePattern) {
                $tempDirs = @($userDirs | Where-Object { $_.Name -match $tempProfilePattern })
            }
            else {
                $tempDirs = @()
            }

            foreach ($user in $tempDirs) {
                $recentFolders = @(
                    Get-ChildItem $user.FullName -Force -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -in $activityFolders } |
                    Where-Object { $_.LastWriteTime -gt $tempCutDate }
                )

                if ($recentFolders.Count -lt 1) {
                    $didRemove = Remove-UserProfile -UserDir $user
                    if ($didRemove) {
                        $tempRemoved += $user.FullName
                    }
                }
            }


            # --- Phase 4: Clean orphaned ProfileList registry entries ---
            $orphanRemoved = @()
            $plKeys = @(Get-ChildItem $regProfiles -ErrorAction SilentlyContinue)

            foreach ($key in $plKeys) {
                # Skip .bak keys -- handled in Phase 5
                if ($key.PSChildName -match '\.bak$') { continue }

                $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                if (-not $props.ProfileImagePath) { continue }

                $profilePath = $props.ProfileImagePath

                # Skip system profiles
                if ($profilePath -match "\\(systemprofile|LocalService|NetworkService)$") {
                    continue
                }

                # If the folder no longer exists, this is an orphan
                if (-not (Test-Path $profilePath)) {
                    $sid = $key.PSChildName

                    if (-not $dryRun) {
                        Remove-Item $key.PSPath -Force -Recurse -ErrorAction SilentlyContinue

                        # Clean corresponding ProfileGuid entry
                        if (Test-Path $regProfileGuid) {
                            $guidKeys = @(Get-ChildItem $regProfileGuid -ErrorAction SilentlyContinue)
                            foreach ($gk in $guidKeys) {
                                $gkProps = Get-ItemProperty $gk.PSPath -ErrorAction SilentlyContinue
                                if ($gkProps.SidString -eq $sid) {
                                    Remove-Item $gk.PSPath -Force -Recurse -ErrorAction SilentlyContinue
                                }
                            }
                        }

                        # Clean Group Policy cache for this SID
                        $gpSidPath = Join-Path $regGPState $sid
                        if (Test-Path $gpSidPath) {
                            Remove-Item $gpSidPath -Force -Recurse -ErrorAction SilentlyContinue
                        }
                    }

                    $orphanRemoved += Resolve-SidToName -Sid $sid
                }
            }


            # --- Phase 5: Clean .bak ProfileList entries ---
            # When Windows fails to load a profile, it renames the registry key
            # from S-1-5-21-xxx to S-1-5-21-xxx.bak. These stale entries cause
            # "temporary profile" login failures for future logons.
            $bakRemoved = @()
            $plKeys = @(Get-ChildItem $regProfiles -ErrorAction SilentlyContinue)

            foreach ($key in $plKeys) {
                if ($key.PSChildName -notmatch '\.bak$') { continue }

                $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                $profilePath = $props.ProfileImagePath

                # Only remove .bak entries whose profile folder is gone or whose
                # real (non-.bak) SID key already exists (duplicate from a failed load)
                $realSid = $key.PSChildName -replace '\.bak$', ''
                $realKeyPath = Join-Path $regProfiles $realSid
                $realKeyExists = Test-Path $realKeyPath
                $folderExists = $false
                if ($profilePath) {
                    $folderExists = Test-Path $profilePath
                }

                if ($realKeyExists -or -not $folderExists) {
                    if (-not $dryRun) {
                        Remove-Item $key.PSPath -Force -Recurse -ErrorAction SilentlyContinue
                    }
                    $bakRemoved += Resolve-SidToName -Sid $realSid
                }
            }


            # --- Phase 6: Clean orphaned ProfileGuid entries ---
            # ProfileGuid maps {GUID} -> SID. If the SID no longer has a
            # ProfileList entry, the GUID mapping is stale.
            $guidRemoved = @()

            if (Test-Path $regProfileGuid) {
                # Build set of valid SIDs from current ProfileList
                $validSids = @{}
                $plKeys = @(Get-ChildItem $regProfiles -ErrorAction SilentlyContinue)
                foreach ($key in $plKeys) {
                    $validSids[$key.PSChildName] = $true
                }

                $guidKeys = @(Get-ChildItem $regProfileGuid -ErrorAction SilentlyContinue)
                foreach ($gk in $guidKeys) {
                    $gkProps = Get-ItemProperty $gk.PSPath -ErrorAction SilentlyContinue
                    $sidString = $gkProps.SidString
                    if (-not $sidString) { continue }

                    if (-not $validSids.ContainsKey($sidString)) {
                        if (-not $dryRun) {
                            Remove-Item $gk.PSPath -Force -Recurse -ErrorAction SilentlyContinue
                        }
                        $guidRemoved += Resolve-SidToName -Sid $sidString
                    }
                }
            }


            # --- Phase 7: Clean orphaned Group Policy State entries ---
            # Per-user GP cache under HKLM\...\Group Policy\State\<SID>.
            # Rebuilt automatically on next logon; safe to remove for deleted users.
            $gpRemoved = @()

            if (Test-Path $regGPState) {
                # Rebuild valid SIDs list (ProfileList may have changed after cleanup)
                $validSids = @{}
                $plKeys = @(Get-ChildItem $regProfiles -ErrorAction SilentlyContinue)
                foreach ($key in $plKeys) {
                    $sidName = $key.PSChildName -replace '\.bak$', ''
                    $validSids[$sidName] = $true
                }

                $gpKeys = @(Get-ChildItem $regGPState -ErrorAction SilentlyContinue)
                foreach ($gpKey in $gpKeys) {
                    $gpSid = $gpKey.PSChildName

                    # Skip well-known system SIDs (S-1-5-18 = SYSTEM, S-1-5-19 = LOCAL SERVICE, etc.)
                    if ($gpSid -match '^S-1-5-(18|19|20)$') { continue }

                    if (-not $validSids.ContainsKey($gpSid)) {
                        if (-not $dryRun) {
                            Remove-Item $gpKey.PSPath -Force -Recurse -ErrorAction SilentlyContinue
                        }
                        $gpRemoved += Resolve-SidToName -Sid $gpSid
                    }
                }
            }


            # --- Calculate results ---
            $totalRemoved = $oldRemoved.Count + $emptyRemoved.Count + $tempRemoved.Count
            $totalRegCleanup = $orphanRemoved.Count + $bakRemoved.Count + $guidRemoved.Count + $gpRemoved.Count

            if ($dryRun) {
                if ($totalRemoved -eq 0 -and $totalRegCleanup -eq 0) {
                    $comment = "No profiles eligible for deletion"
                }
                else {
                    $comment = "WhatIf -- no changes made"
                }
            }
            else {
                if ($totalRemoved -eq 0 -and $totalRegCleanup -eq 0) {
                    $comment = "No profiles eligible for deletion"
                }
                else {
                    $comment = "Cleanup complete"
                }
            }

            # Build list of users that were protected from deletion
            $skippedUsers = @()
            foreach ($sessionUser in $activeSessionUsers.Keys) {
                $skippedUsers += "$sessionUser (active session)"
            }
            foreach ($loadedPath in $loadedPaths.Keys) {
                $loadedName = Split-Path $loadedPath -Leaf
                if (-not $activeSessionUsers.ContainsKey($loadedName)) {
                    $skippedUsers += "$loadedName (loaded profile)"
                }
            }
            foreach ($specialPath in $specialPaths.Keys) {
                $specialName = Split-Path $specialPath -Leaf
                $skippedUsers += "$specialName (special/system)"
            }
            foreach ($protectedPath in $protectedStatePaths.Keys) {
                $protectedName = Split-Path $protectedPath -Leaf
                $skippedUsers += "$protectedName (mandatory/creating)"
            }
            foreach ($svcUser in $serviceAccountUsers.Keys) {
                $reason = $serviceAccountUsers[$svcUser]
                $skippedUsers += "$svcUser ($reason)"
            }

            # Extract just the usernames from full paths for readability
            $oldNames    = @($oldRemoved   | ForEach-Object { Split-Path $_ -Leaf })
            $emptyNames  = @($emptyRemoved | ForEach-Object { Split-Path $_ -Leaf })
            $tempNames   = @($tempRemoved  | ForEach-Object { Split-Path $_ -Leaf })

            $cleanDisk = Get-PSDrive ($env:SystemDrive).Replace(":", "")
            $cleanFreeGB = [math]::Round($($cleanDisk.Free / 1gb), 2)

            if ($dryRun) {
                $freedSizeGB = $null
            }
            else {
                $freedSizeGB = [math]::Round($cleanFreeGB - $diskFreeGB, 2)
            }

            [PSCustomObject]@{
                OldRemoved    = $oldRemoved.Count
                EmptyRemoved  = $emptyRemoved.Count
                TempRemoved   = $tempRemoved.Count
                OrphanRemoved = $orphanRemoved.Count
                BakRemoved    = $bakRemoved.Count
                GuidRemoved   = $guidRemoved.Count
                GPRemoved     = $gpRemoved.Count
                OldNames      = $oldNames
                EmptyNames    = $emptyNames
                TempNames     = $tempNames
                OrphanNames   = $orphanRemoved
                BakNames      = $bakRemoved
                GuidNames     = $guidRemoved
                GPNames       = $gpRemoved
                SkippedUsers  = $skippedUsers
                SpaceFreed    = $freedSizeGB
                TotalFree     = $cleanFreeGB
                Comment       = $comment
            }
        }

        $result = Invoke-Command -ComputerName $computer -ScriptBlock $remoteBlock `
            -ArgumentList $size, $userAge, $tempAge, $dryRun

        # Transform for output
        if ($result) {
            [PSCustomObject]@{
                ComputerName     = $computer
                "Old Profiles"   = $result.OldRemoved
                "Empty Profiles" = $result.EmptyRemoved
                "Temp Profiles"  = $result.TempRemoved
                "Orphan Keys"    = $result.OrphanRemoved
                "Bak Keys"       = $result.BakRemoved
                "Guid Keys"      = $result.GuidRemoved
                "GP Keys"        = $result.GPRemoved
                OldNames         = $result.OldNames
                EmptyNames       = $result.EmptyNames
                TempNames        = $result.TempNames
                OrphanNames      = $result.OrphanNames
                BakNames         = $result.BakNames
                GuidNames        = $result.GuidNames
                GPNames          = $result.GPNames
                SkippedUsers     = $result.SkippedUsers
                "Space Freed"    = $result.SpaceFreed
                "Total Free"     = $result.TotalFree
                Comment          = $result.Comment
            }
        }
    }


    # --- Test connection ---
    Write-Host ""
    Write-Host "Checking for online machines..."

    $pingResults = Test-ConnectionAsJob -ComputerName $list

    $offlineResults = @()
    $onlineList = @()

    foreach ($pingResult in $pingResults) {
        if ($pingResult.Reachable -eq $true) {
            $onlineList += $pingResult.ComputerName
        }
        else {
            $offlineResults += [PSCustomObject]@{
                ComputerName     = $pingResult.ComputerName
                "Old Profiles"   = $null
                "Empty Profiles" = $null
                "Temp Profiles"  = $null
                "Orphan Keys"    = $null
                "Bak Keys"       = $null
                "Guid Keys"      = $null
                "GP Keys"        = $null
                OldNames         = @()
                EmptyNames       = @()
                TempNames        = @()
                OrphanNames      = @()
                BakNames         = @()
                GuidNames        = @()
                GPNames          = @()
                SkippedUsers     = @()
                "Space Freed"    = $null
                "Total Free"     = $null
                Comment          = "Offline"
            }
        }
    }

    if ($onlineList.Count -eq 0) {
        $offlineSorted = @($offlineResults | Sort-Object ComputerName)
        $offlineSorted | Format-Table -AutoSize | Out-Host
        if ($PassThru) { return $offlineSorted }
        return
    }


    # --- Run cleanup via Invoke-RunspacePool ---
    $argList = foreach ($computer in $onlineList) {
        , @($computer, $MinFreeGB, $UserMaxAgeDays, $TempMaxAgeDays, $dryRun)
    }

    $activityLabel = "Cleanup C:\Users"
    if ($dryRun) { $activityLabel = "Cleanup C:\Users (WhatIf)" }

    $results = @(Invoke-RunspacePool -ScriptBlock $scriptblock -ArgumentList $argList `
        -ThrottleLimit $ThrottleLimit -TimeoutMinutes $TimeoutMinutes `
        -ActivityName $activityLabel)

    # Normalize timed-out/failed results from RunspacePool into the same shape
    $allResults = @($results) + @($offlineResults)

    $finalResults = foreach ($r in $allResults) {
        if ($r.PSObject.Properties.Name -contains "Old Profiles") {
            $r
        }
        else {
            [PSCustomObject]@{
                ComputerName     = $r.ComputerName
                "Old Profiles"   = $null
                "Empty Profiles" = $null
                "Temp Profiles"  = $null
                "Orphan Keys"    = $null
                "Bak Keys"       = $null
                "Guid Keys"      = $null
                "GP Keys"        = $null
                OldNames         = @()
                EmptyNames       = @()
                TempNames        = @()
                OrphanNames      = @()
                BakNames         = @()
                GuidNames        = @()
                GPNames          = @()
                SkippedUsers     = @()
                "Space Freed"    = $null
                "Total Free"     = $null
                Comment          = $r.Comment
            }
        }
    }


    # --- Output results ---

    $sortedResults = @($finalResults | Sort-Object -Property (
        @{Expression = { if ($_.Comment -eq "Offline") { 1 } else { 0 } }},
        @{Expression = "Total Free"; Descending = $false},
        @{Expression = "Space Freed"; Descending = $false},
        @{Expression = "Comment"}
    ))

    # --- Header ---
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan

    if ($dryRun) {
        Write-Host "  Cleanup C:\Users -- WhatIf Preview" -ForegroundColor Cyan
        Write-Host "  (No changes were made)" -ForegroundColor Yellow
    }
    else {
        Write-Host "  Cleanup C:\Users -- Results Summary" -ForegroundColor Cyan
    }

    Write-Host "========================================" -ForegroundColor Cyan

    if ($dryRun) {
        Write-Host ""
        Write-Host "  Parameters:" -ForegroundColor Gray
        Write-Host "    MinFreeGB:      $MinFreeGB GB" -ForegroundColor Gray
        Write-Host "    UserMaxAgeDays: $UserMaxAgeDays days" -ForegroundColor Gray
        Write-Host "    TempMaxAgeDays: $TempMaxAgeDays days" -ForegroundColor Gray
    }

    Write-Host ""


    # --- Summary table ---
    # Profile removal columns
    $tableProps = @(
        "ComputerName"
        "Old Profiles"
        "Empty Profiles"
        "Temp Profiles"
    )

    if (-not $dryRun) {
        $tableProps += "Space Freed"
    }

    $tableProps += "Total Free"
    $tableProps += "Comment"

    $sortedResults | Select-Object $tableProps | Format-Table -AutoSize | Out-Host


    # --- Registry cleanup table (only if there was registry work) ---
    $hasRegWork = $false
    foreach ($r in $sortedResults) {
        if ($r."Orphan Keys" -gt 0 -or $r."Bak Keys" -gt 0 -or
            $r."Guid Keys" -gt 0 -or $r."GP Keys" -gt 0) {
            $hasRegWork = $true
            break
        }
    }

    if ($hasRegWork) {
        Write-Host "  Registry Cleanup:" -ForegroundColor Cyan
        Write-Host ""

        $regTableProps = @(
            "ComputerName"
            "Orphan Keys"
            "Bak Keys"
            "Guid Keys"
            "GP Keys"
        )

        $regResults = @($sortedResults | Where-Object {
            $_."Orphan Keys" -gt 0 -or $_."Bak Keys" -gt 0 -or
            $_."Guid Keys" -gt 0 -or $_."GP Keys" -gt 0
        })

        $regResults | Select-Object $regTableProps | Format-Table -AutoSize | Out-Host
    }


    # --- Per-machine detail ---
    $detailLabel = if ($dryRun) { "would be removed" } else { "removed" }

    $detailResults = @($sortedResults | Where-Object {
        $_.Comment -match "(Cleanup complete|WhatIf)" -or
        @($_.SkippedUsers).Count -gt 0
    })

    if ($detailResults.Count -gt 0) {
        Write-Host "----------------------------------------" -ForegroundColor DarkGray
        Write-Host "  Per-Machine Detail" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor DarkGray

        foreach ($r in $detailResults) {
            Write-Host ""
            Write-Host "  $($r.ComputerName)" -ForegroundColor Yellow

            if (@($r.OldNames).Count -gt 0) {
                Write-Host "    Old profiles ($detailLabel):   " -ForegroundColor Gray -NoNewline
                Write-Host ($r.OldNames -join ", ")
            }
            if (@($r.EmptyNames).Count -gt 0) {
                Write-Host "    Empty profiles ($detailLabel): " -ForegroundColor Gray -NoNewline
                Write-Host ($r.EmptyNames -join ", ")
            }
            if (@($r.TempNames).Count -gt 0) {
                Write-Host "    Temp profiles ($detailLabel):  " -ForegroundColor Gray -NoNewline
                Write-Host ($r.TempNames -join ", ")
            }
            if (@($r.OrphanNames).Count -gt 0) {
                Write-Host "    Orphan keys ($detailLabel):    " -ForegroundColor Gray -NoNewline
                Write-Host ($r.OrphanNames -join ", ")
            }
            if (@($r.BakNames).Count -gt 0) {
                Write-Host "    Bak keys ($detailLabel):       " -ForegroundColor Gray -NoNewline
                Write-Host ($r.BakNames -join ", ")
            }
            if (@($r.GuidNames).Count -gt 0) {
                Write-Host "    Guid keys ($detailLabel):      " -ForegroundColor Gray -NoNewline
                Write-Host ($r.GuidNames -join ", ")
            }
            if (@($r.GPNames).Count -gt 0) {
                Write-Host "    GP cache ($detailLabel):       " -ForegroundColor Gray -NoNewline
                Write-Host ($r.GPNames -join ", ")
            }
            if (@($r.SkippedUsers).Count -gt 0) {
                Write-Host "    Protected (skipped):           " -ForegroundColor Gray -NoNewline
                Write-Host ($r.SkippedUsers -join ", ") -ForegroundColor DarkYellow
            }
        }
        Write-Host ""
    }


    # --- Totals ---
    $totalOld     = ($sortedResults | ForEach-Object { $_."Old Profiles" }   | Measure-Object -Sum).Sum
    $totalEmpty   = ($sortedResults | ForEach-Object { $_."Empty Profiles" } | Measure-Object -Sum).Sum
    $totalTemp    = ($sortedResults | ForEach-Object { $_."Temp Profiles" }  | Measure-Object -Sum).Sum
    $totalOrphan  = ($sortedResults | ForEach-Object { $_."Orphan Keys" }    | Measure-Object -Sum).Sum
    $totalBak     = ($sortedResults | ForEach-Object { $_."Bak Keys" }       | Measure-Object -Sum).Sum
    $totalGuid    = ($sortedResults | ForEach-Object { $_."Guid Keys" }      | Measure-Object -Sum).Sum
    $totalGP      = ($sortedResults | ForEach-Object { $_."GP Keys" }        | Measure-Object -Sum).Sum
    $totalProfiles = $totalOld + $totalEmpty + $totalTemp
    $totalRegKeys  = $totalOrphan + $totalBak + $totalGuid + $totalGP

    $machinesOnline  = @($sortedResults | Where-Object { $_.Comment -ne "Offline" }).Count
    $machinesOffline = @($sortedResults | Where-Object { $_.Comment -eq "Offline" }).Count
    $machinesCleaned = @($sortedResults | Where-Object {
        $_.Comment -eq "Cleanup complete" -or $_.Comment -eq "WhatIf -- no changes made"
    }).Count
    $machinesSkipped = @($sortedResults | Where-Object { $_.Comment -eq "Cleanup not initiated" }).Count
    $machinesNone    = @($sortedResults | Where-Object { $_.Comment -eq "No profiles eligible for deletion" }).Count
    $machinesFailed  = @($sortedResults | Where-Object {
        $_.Comment -match "Task (Stopped|Failed|Error)" -or $_.Comment -eq "Job Failed"
    }).Count

    $profileVerb = if ($dryRun) { "flagged for removal" } else { "removed" }

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Totals" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Machines targeted:    $($sortedResults.Count)" -ForegroundColor Gray
    Write-Host "    Online:             $machinesOnline" -ForegroundColor Gray
    Write-Host "    Offline:            $machinesOffline" -ForegroundColor Gray
    Write-Host ""

    if ($dryRun) {
        Write-Host "    Would be cleaned:   $machinesCleaned" -ForegroundColor Yellow
    }
    else {
        Write-Host "    Cleaned:            $machinesCleaned" -ForegroundColor Green
    }

    Write-Host "    Disk OK (>$MinFreeGB GB):   $machinesSkipped" -ForegroundColor Gray
    Write-Host "    Nothing to clean:   $machinesNone" -ForegroundColor Gray

    if ($machinesFailed -gt 0) {
        Write-Host "    Failed/Stopped:     $machinesFailed" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "  Profiles $($profileVerb):  $totalProfiles" -ForegroundColor Gray
    Write-Host "    Old:                $totalOld" -ForegroundColor Gray
    Write-Host "    Empty:              $totalEmpty" -ForegroundColor Gray
    Write-Host "    Temp:               $totalTemp" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Registry keys $($profileVerb): $totalRegKeys" -ForegroundColor Gray
    Write-Host "    Orphan ProfileList: $totalOrphan" -ForegroundColor Gray
    Write-Host "    Bak ProfileList:    $totalBak" -ForegroundColor Gray
    Write-Host "    Orphan ProfileGuid: $totalGuid" -ForegroundColor Gray
    Write-Host "    GP State cache:     $totalGP" -ForegroundColor Gray

    if (-not $dryRun) {
        $totalFreed   = ($sortedResults | ForEach-Object { $_."Space Freed" } | Measure-Object -Sum).Sum
        $totalFreedGB = [math]::Round($totalFreed, 2)
        Write-Host ""
        Write-Host "  Total space freed:    $totalFreedGB GB" -ForegroundColor Gray
    }

    Write-Host ""

    if ($PassThru) { return $sortedResults }

    } # end 'end' block
}
