$desktopPath = [Environment]::GetFolderPath("Desktop")
$LogFile = Join-Path -Path $desktopPath -ChildPath "OneDrive_Migration.log"
$script:PollSeconds = 2
$script:HydrateTimeoutSeconds = 1800
$script:MaxPathLength = 260

Set-StrictMode -Version Latest

$script:OriginalProgressPreference = $ProgressPreference
$ProgressPreference = 'Continue'

# --- Files to skip ---
$SkipFileNames = @('desktop.ini', 'Thumbs.db', '.ds_store', '.sync', '.owncloudsync.log')

# ===================== Helper functions =====================

$script:ConsoleLogEnabled = $true

function Write-Log {
    param(
        [string]$Text,
        [string]$Level = "INFO",
        [switch]$ForceConsole
    )
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Text"
    if ($ForceConsole -or $script:ConsoleLogEnabled -or $Level -in @("ERROR", "WARN")) {
        Write-Information $line -InformationAction Continue
    }
    if ($LogFile) {
        $logDir = Split-Path -Path $LogFile -Parent
        if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        $line | Out-File -FilePath $LogFile -Append -Encoding UTF8
    }
}

function Update-ProgressUi {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [int]$Index,
        [int]$Total,
        [string]$CurrentTask,
        [string]$LastMoved
    )

    if (-not $PSCmdlet.ShouldProcess("Console", "Update progress UI")) { return }

    $percent = 0
    if ($Total -gt 0) {
        $percent = [Math]::Min([Math]::Floor(($Index / $Total) * 100), 100)
    }

    if ([string]::IsNullOrWhiteSpace($CurrentTask)) { $CurrentTask = "Ready..." }
    if ([string]::IsNullOrWhiteSpace($LastMoved)) { $LastMoved = "None yet" }

    $status = "{0}%  |  {1}/{2} files" -f $percent, $Index, $Total
    $currentOp = "{0} | Last: {1}" -f $CurrentTask, $LastMoved
    Write-Progress -Activity "OneDrive Migration" -Status $status -CurrentOperation $currentOp -PercentComplete $percent
}

function Confirm-YesNo {
    param(
        [string]$Message,
        [bool]$DefaultYes = $true
    )

    $hint = if ($DefaultYes) { "Y/n" } else { "y/N" }
    while ($true) {
        $response = Read-Input "$Message [$hint]"
        if ([string]::IsNullOrWhiteSpace($response)) { return $DefaultYes }
        if ($response -match '^(y|yes)$') { return $true }
        if ($response -match '^(n|no)$') { return $false }
        Write-Information "Please answer Y or N." -InformationAction Continue
    }
}

function Read-Input {
    param([string]$Prompt)
    return (Read-Host $Prompt)
}

function Show-Header {
    $lines = @(
        "========================================",
        "          OneDrive Migration",
        "  Copy OneDrive files to Nextcloud or",
        "      a custom destination folder",
        '  All files go into "OneDriveMigration"',
        "  with verification and optional source dehydrate",
        "  Log file will be saved on your Desktop",
        "========================================"
    )

    foreach ($line in $lines) {
        Write-Information $line -InformationAction Continue
    }
}

function Get-SourceDehydrateMode {
    Write-Information "Dehydrate = mark file as cloud-only to free local disk space." -InformationAction Continue
    Write-Information "Choose source dehydrate mode:" -InformationAction Continue
    Write-Information "  1) Only files hydrated during this run (default)" -InformationAction Continue
    Write-Information "  2) All source files" -InformationAction Continue
    Write-Information "  3) Do not dehydrate source files" -InformationAction Continue

    while ($true) {
        $choice = Read-Input "Select 1, 2, or 3"
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice -eq "1") { return "HydratedOnly" }
        if ($choice -eq "2") { return "All" }
        if ($choice -eq "3") { return "None" }
        Write-Information "Invalid selection." -InformationAction Continue
    }
}

function Get-DestinationMode {
    Write-Information "Choose destination mode:" -InformationAction Continue
    Write-Information "  1) Nextcloud (default)" -InformationAction Continue
    Write-Information "  2) Custom destination folder" -InformationAction Continue

    while ($true) {
        $choice = Read-Input "Select 1 or 2"
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice -eq "1") { return "Nextcloud" }
        if ($choice -eq "2") { return "Custom" }
        Write-Information "Invalid selection." -InformationAction Continue
    }
}

function Get-CloudDestination {
    $candidates = @(
        [pscustomobject]@{ Name = "Nextcloud"; Path = (Get-NextcloudRoot) },
        [pscustomobject]@{ Name = "Dropbox"; Path = $env:DROPBOX },
        [pscustomobject]@{ Name = "Dropbox"; Path = $env:Dropbox },
        [pscustomobject]@{ Name = "Google Drive"; Path = (Join-Path -Path $env:USERPROFILE -ChildPath "Google Drive") },
        [pscustomobject]@{ Name = "Google Drive"; Path = (Join-Path -Path $env:USERPROFILE -ChildPath "Google Drive (Personal)") },
        [pscustomobject]@{ Name = "Google Drive"; Path = (Join-Path -Path $env:USERPROFILE -ChildPath "Google Drive (Work)") },
        [pscustomobject]@{ Name = "iCloud Drive"; Path = $env:ICLOUDDRIVE },
        [pscustomobject]@{ Name = "iCloud Drive"; Path = (Join-Path -Path $env:USERPROFILE -ChildPath "iCloud Drive") },
        [pscustomobject]@{ Name = "iCloud Drive"; Path = (Join-Path -Path $env:USERPROFILE -ChildPath "iCloudDrive") },
        [pscustomobject]@{ Name = "iCloud Drive"; Path = (Join-Path -Path $env:USERPROFILE -ChildPath "iCloudDrive\iCloud Drive") },
        [pscustomobject]@{ Name = "Box"; Path = (Join-Path -Path $env:USERPROFILE -ChildPath "Box") },
        [pscustomobject]@{ Name = "MEGA"; Path = (Join-Path -Path $env:USERPROFILE -ChildPath "MEGA") }
    )

    $detected = @()
    foreach ($item in $candidates) {
        if (-not $item.Path) { continue }
        if (Test-Path -LiteralPath $item.Path) {
            $detected += $item
        }
    }

    return ($detected | Sort-Object Name, Path -Unique)
}

function Select-DestinationRoot {
    $destinations = @(Get-CloudDestination)
    $options = @()

    if ($destinations.Count -gt 0) {
        Write-Information "Detected destination folders:" -InformationAction Continue
        for ($i = 0; $i -lt $destinations.Count; $i++) {
            $label = "{0} ({1})" -f $destinations[$i].Name, $destinations[$i].Path
            $options += [pscustomobject]@{ Index = $i + 1; Name = $destinations[$i].Name; Path = $destinations[$i].Path; Label = $label }
            Write-Information ("  {0}) {1}" -f ($i + 1), $label) -InformationAction Continue
        }
        Write-Information ("  {0}) Custom destination folder" -f ($destinations.Count + 1)) -InformationAction Continue
    } else {
        Write-Information "No known cloud destinations detected." -InformationAction Continue
        Write-Information "  1) Custom destination folder" -InformationAction Continue
    }

    while ($true) {
        $choice = Read-Input "Select a destination"

        if ($destinations.Count -eq 0) {
            if ($choice -eq "1") { return [pscustomobject]@{ Mode = "Custom"; Root = $null; Label = "Custom" } }
            Write-Information "Invalid selection." -InformationAction Continue
            continue
        }

        if ($choice -match '^\d+$') {
            $idx = [int]$choice
            if ($idx -ge 1 -and $idx -le $destinations.Count) {
                $selected = $destinations[$idx - 1]
                Write-Information "" -InformationAction Continue
                return [pscustomobject]@{ Mode = $selected.Name; Root = $selected.Path; Label = $selected.Name }
            }
            if ($idx -eq ($destinations.Count + 1)) {
                Write-Information "" -InformationAction Continue
                return [pscustomobject]@{ Mode = "Custom"; Root = $null; Label = "Custom" }
            }
        }

        Write-Information "Invalid selection." -InformationAction Continue
    }
}

function Test-DehydrateSource {
    param(
        [string]$Mode,
        [bool]$WasHydrated
    )

    switch ($Mode) {
        "All" { return $true }
        "None" { return $false }
        default { return $WasHydrated }
    }
}

function Get-OneDriveAccount {
    $accounts = @()
    $accountsKey = "HKCU:\Software\Microsoft\OneDrive\Accounts"

    if (Test-Path -LiteralPath $accountsKey) {
        Get-ChildItem -LiteralPath $accountsKey | ForEach-Object {
            $p = Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue
            if (-not $p) { return }

            $userFolderProp = $p.PSObject.Properties["UserFolder"]
            if (-not $userFolderProp) { return }

            $accountTypeProp = $p.PSObject.Properties["AccountType"]
            $accountTypeValue = if ($accountTypeProp) { $accountTypeProp.Value } else { $null }
            if (-not $accountTypeValue) {
                if ($_.PSChildName -match 'Business|Commercial|Enterprise' -or $userFolderProp.Value -match 'OneDrive\s-\s') {
                    $accountTypeValue = "Business"
                } elseif ($_.PSChildName -match 'Personal|Default' -or $userFolderProp.Value -match 'OneDrive$') {
                    $accountTypeValue = "Personal"
                }
            }

            $accounts += [pscustomobject]@{
                Name        = $_.PSChildName
                UserFolder  = $userFolderProp.Value
                AccountType = $accountTypeValue
            }
        }
    }

    if ($accounts.Count -eq 0 -and $env:OneDrive) {
        $accounts += [pscustomobject]@{
            Name        = "Default"
            UserFolder  = $env:OneDrive
            AccountType = "Personal"
        }
    }

    return $accounts
}

function Get-OneDriveAccountType {
    param([object[]]$Accounts)

    if ($Accounts | Where-Object { $_.AccountType -match 'Business' }) { return "Business" }
    if ($Accounts | Where-Object { $_.AccountType -match 'Personal' }) { return "Personal" }
    return "Unknown"
}

function Select-OneDriveFolder {
    param([string[]]$Folders)

    if ($Folders.Count -gt 0) {
        Write-Information "Detected OneDrive folders:" -InformationAction Continue
        for ($i = 0; $i -lt $Folders.Count; $i++) {
            Write-Information "  $($i + 1)) $($Folders[$i])" -InformationAction Continue
        }
        Write-Information "  2) Custom path" -InformationAction Continue

        while ($true) {
            $choice = Read-Input "Select folder"
            if ($choice -match '[:\\/]') {
                if (Test-Path -LiteralPath $choice) {
                    Write-Information "" -InformationAction Continue
                    return $choice
                }
                Write-Information "Path not found: $choice" -InformationAction Continue
                continue
            }
            if ($choice -eq "2") {
                $custom = Read-Input "Enter custom OneDrive path"
                if (Test-Path -LiteralPath $custom) {
                    Write-Information "" -InformationAction Continue
                    return $custom
                }
                Write-Information "Path not found: $custom" -InformationAction Continue
                continue
            }

            if ($choice -match '^\d+$') {
                $idx = [int]$choice
                if ($idx -ge 1 -and $idx -le $Folders.Count) {
                    $selected = $Folders[$idx - 1]
                    if (Confirm-YesNo -Message "Use this folder? $selected" -DefaultYes $true) {
                        Write-Information "" -InformationAction Continue
                        return $selected
                    }
                    continue
                }
            }

            Write-Information "Invalid selection." -InformationAction Continue
        }
    }

    while ($true) {
        $custom = Read-Input "Enter OneDrive path"
        if (Test-Path -LiteralPath $custom) {
            Write-Information "" -InformationAction Continue
            return $custom
        }
        Write-Information "Path not found: $custom" -InformationAction Continue
    }
}

function Get-NextcloudRoot {
    $candidates = @()
    if ($env:NEXTCLOUD) { $candidates += $env:NEXTCLOUD }
    if ($env:Nextcloud) { $candidates += $env:Nextcloud }
    if ($env:USERPROFILE) { $candidates += (Join-Path -Path $env:USERPROFILE -ChildPath "Nextcloud") }

    $candidates = $candidates | Where-Object { $_ } | Select-Object -Unique
    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) { return $path }
    }

    return $null
}

function New-FolderIfMissing {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        if ($PSCmdlet.ShouldProcess($Path, "Create directory")) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
    }
}

function Test-CloudOnly {
    param([string]$Path)
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        # Offline attribute = file is cloud-only (not stored locally)
        return (($item.Attributes -band [IO.FileAttributes]::Offline) -ne 0)
    } catch {
        Write-Log "Attribute check failed: $Path | $_" -Level "WARN"
        return $false
    }
}

function Test-PathTooLong {
    param(
        [string]$Path,
        [int]$MaxLength = $script:MaxPathLength
    )

    if ([string]::IsNullOrEmpty($Path)) { return $false }
    return ($Path.Length -ge $MaxLength)
}

function Invoke-HydrateFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "File not found: $Path"
    }

    if (-not (Test-CloudOnly -Path $Path)) {
        return $false
    }

    # Set pin attribute -> OneDrive downloads the file
    $attribOut = & attrib.exe +P -U "$Path" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "attrib +P -U failed ($LASTEXITCODE): $attribOut"
    }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $lastErr = $null
    # Progressive polling: check fast first (200ms), then slow down
    # Avoids unnecessary waiting for small files
    $currentWaitMs = 200

    while ($true) {
        # Try to read the file - only when this works is it actually local
        try {
            $fs = [System.IO.File]::Open(
                $Path,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite
            )
            try {
                $buf = New-Object byte[] 1
                [void]$fs.Read($buf, 0, 1)
            } finally {
                $fs.Close()
                $fs.Dispose()
            }
        } catch {
            $lastErr = $_.Exception.Message
        }

        if (-not (Test-CloudOnly -Path $Path)) { break }

        if ($sw.Elapsed.TotalSeconds -ge $script:HydrateTimeoutSeconds) {
            $msg = "Hydrate timeout ($script:HydrateTimeoutSeconds s): $Path"
            if ($lastErr) { $msg += " | LastError: $lastErr" }
            throw $msg
        }

        Start-Sleep -Milliseconds $currentWaitMs
        # Increase progressively: 200ms -> 400ms -> 800ms -> 1600ms -> max 2000ms
        $currentWaitMs = [Math]::Min($currentWaitMs * 2, $script:PollSeconds * 1000)
    }

    return $true
}

function Invoke-DehydrateFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return }

    if (Test-CloudOnly -Path $Path) {
        return
    }
    $attribOut = & attrib.exe -P +U "$Path" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Log "attrib -P +U failed ($LASTEXITCODE): $attribOut | $Path" -Level "WARN"
        # No throw - not critical, file stays local
    }
}

function Get-RelativePathSafe {
    param([string]$Root, [string]$FullPath)

    $rootNorm = [System.IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    $fullNorm = [System.IO.Path]::GetFullPath($FullPath)

    if (-not $fullNorm.StartsWith($rootNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is not under root. Root=$Root | Path=$FullPath"
    }

    return $fullNorm.Substring($rootNorm.Length)
}

function Copy-FileStreaming {
    param([string]$Src, [string]$Dst)

    $srcStream = $null
    $dstStream = $null

    try {
        $srcStream = [System.IO.File]::Open(
            $Src,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )
        $dstStream = [System.IO.File]::Open(
            $Dst,
            [System.IO.FileMode]::Create,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None
        )

        $buffer = New-Object byte[] (4MB)
        while ($true) {
            $read = $srcStream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            $dstStream.Write($buffer, 0, $read)
        }
        $dstStream.Flush()
    } finally {
        if ($dstStream) { $dstStream.Close(); $dstStream.Dispose() }
        if ($srcStream) { $srcStream.Close(); $srcStream.Dispose() }
    }
}

function Test-Copy {
    param([string]$Src, [string]$Dst)

    if (-not (Test-Path -LiteralPath $Dst)) {
        throw "Destination file does not exist after copy: $Dst"
    }

    $srcSize = (Get-Item -LiteralPath $Src -Force).Length
    $dstSize = (Get-Item -LiteralPath $Dst -Force).Length

    if ($srcSize -ne $dstSize) {
        # Delete destination so no corrupt file remains
        Remove-Item -LiteralPath $Dst -Force -ErrorAction SilentlyContinue
        throw "Size check failed! Src=$srcSize Dst=$dstSize | $Dst"
    }
}

# ===================== Interactive setup =====================

Show-Header

$accounts = Get-OneDriveAccount
$detectedType = Get-OneDriveAccountType -Accounts $accounts
if ($detectedType -eq "Unknown") {
    Write-Information "Detected OneDrive account type: Unknown" -InformationAction Continue
    Write-Information "Defaulting to Personal (can be changed below)." -InformationAction Continue
    $detectedType = "Personal"
} else {
    Write-Information "Detected OneDrive account type: $detectedType" -InformationAction Continue
}

if (-not (Confirm-YesNo -Message "Is this correct?" -DefaultYes $true)) {
    $typed = Read-Input "Enter account type (P=Personal, B=Business)"
    if ($typed -match '^[bB]') { $detectedType = "Business" } else { $detectedType = "Personal" }
}
Write-Information ("Using account type: {0}" -f $detectedType) -InformationAction Continue
Write-Information "" -InformationAction Continue

$oneDriveFolders = @($accounts | Where-Object { $_.UserFolder } | Select-Object -ExpandProperty UserFolder -Unique)
if ($detectedType -eq "Business" -and $oneDriveFolders.Count -gt 1) {
    Write-Information "Multiple business OneDrive folders detected." -InformationAction Continue
}
$SourceRoot = Select-OneDriveFolder -Folders $oneDriveFolders

$sourceDehydrateMode = Get-SourceDehydrateMode
Write-Information ("Source dehydrate mode: {0}" -f $sourceDehydrateMode) -InformationAction Continue
Write-Information "" -InformationAction Continue

$destinationChoice = Select-DestinationRoot
Write-Information ("Destination mode: {0}" -f $destinationChoice.Label) -InformationAction Continue
Write-Information "" -InformationAction Continue

if ($destinationChoice.Mode -ne "Custom") {
    $DestRoot = Join-Path -Path $destinationChoice.Root -ChildPath "OneDriveMigration"
    New-FolderIfMissing -Path $DestRoot
} else {
    while ($true) {
        $customDest = Read-Input "Enter destination folder"
        if (Test-Path -LiteralPath $customDest) {
            $DestRoot = $customDest
            Write-Information "" -InformationAction Continue
            break
        }
        if (Confirm-YesNo -Message "Path not found. Create it?" -DefaultYes $false) {
            New-FolderIfMissing -Path $customDest
            $DestRoot = $customDest
            break
        }
    }
}

# ===================== Main program =====================

try {
    if (-not (Test-Path -LiteralPath $SourceRoot)) {
        throw "SourceRoot does not exist: $SourceRoot"
    }
    if (-not (Test-Path -LiteralPath $DestRoot)) {
        Write-Log "DestRoot does not exist, creating: $DestRoot"
        New-Item -ItemType Directory -Path $DestRoot -Force | Out-Null
    }

    Write-Log "=========================================="
    Write-Log "Migration started"
    Write-Log "Source:  $SourceRoot"
    Write-Log "Dest:    $DestRoot"
    Write-Log "LogFile: $LogFile"
    Write-Log "=========================================="

    $script:ConsoleLogEnabled = $false

    # List files (metadata only, no download)
    # @() ensures we always get an array (even for 0 or 1 file)
    [array]$files = @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse -Force |
        Where-Object { $SkipFileNames -notcontains $_.Name } |
        Sort-Object -Property Length, FullName)

    $total     = $files.Count
    $index     = 0
    $okCount   = 0
    $skipCount = 0
    $errCount  = 0
    $errList   = @()
    $currentTask = ""
    $lastMoved = ""

    Write-Log "Found files: $total (after filter)"
    Update-ProgressUi -Index 0 -Total $total -CurrentTask "Initializing" -LastMoved $lastMoved

    foreach ($file in $files) {
        $index++
        $src = $file.FullName
        $rel = Get-RelativePathSafe -Root $SourceRoot -FullPath $src
        $dst = Join-Path -Path $DestRoot -ChildPath $rel
        $dstDir = Split-Path -Path $dst -Parent
        $wasHydrated = $false

        if ((Test-PathTooLong -Path $src) -or (Test-PathTooLong -Path $dst)) {
            $errCount++
            $errMsg = "ERROR at [$index/$total] $rel : Path too long (max $script:MaxPathLength). Skipped."
            $errList += $errMsg
            Write-Log $errMsg -Level "ERROR"

            $currentTask = "Error (path too long)"
            Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved
            continue
        }

        $currentTask = "Checking file: $rel"
        Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved

        # Already copied? Compare size (resume support)
        if (Test-Path -LiteralPath $dst) {
            $existingSize = (Get-Item -LiteralPath $dst -Force).Length
            $sourceSize   = $file.Length
            if ($existingSize -eq $sourceSize -and $sourceSize -gt 0) {
                # Dehydrate source if still local
                $currentTask = "Skipped (already exists)"
                Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved
                if (Test-DehydrateSource -Mode $sourceDehydrateMode -WasHydrated $wasHydrated) {
                    Invoke-DehydrateFile -Path $src
                }
                $skipCount++
                continue
            }
        }

        try {
            # 1. Download file (hydrate)
            $currentTask = "Hydrate (download)"
            Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved
            $wasHydrated = Invoke-HydrateFile -Path $src

            # 2. Ensure destination folder
            $currentTask = "Ensure destination folder"
            Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved
            New-FolderIfMissing -Path $dstDir

            # 3. Copy file
            $currentTask = "Copying file"
            Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved
            Copy-FileStreaming -Src $src -Dst $dst

            # 4. Verify copy (size check)
            $currentTask = "Verifying copy"
            Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved
            Test-Copy -Src $src -Dst $dst

            # 5. Dehydrate source (free space on OneDrive)
            $currentTask = "Free space (Source)"
            Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved
            if (Test-DehydrateSource -Mode $sourceDehydrateMode -WasHydrated $wasHydrated) {
                Invoke-DehydrateFile -Path $src
            }

            Write-Log "Completed: $rel"

            $okCount++
            $lastMoved = $rel
            $currentTask = "Done"
            Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved

        } catch {
            $errCount++
            $errMsg = "ERROR at [$index/$total] $rel : $($_.Exception.Message)"
            $errList += $errMsg
            Write-Log $errMsg -Level "ERROR"

            $currentTask = "Error (see log)"
            Update-ProgressUi -Index $index -Total $total -CurrentTask $currentTask -LastMoved $lastMoved

            # Dehydrate source if it was hydrated
            try {
                if (Test-DehydrateSource -Mode $sourceDehydrateMode -WasHydrated $wasHydrated) {
                    Invoke-DehydrateFile -Path $src
                }
            } catch {
                Write-Log "Cleanup dehydrate failed: $src | $_" -Level "WARN"
            }

            # Delete faulty destination file
            if (Test-Path -LiteralPath $dst) {
                Remove-Item -LiteralPath $dst -Force -ErrorAction SilentlyContinue
            }

            # Continue with next file instead of aborting
            continue
        }
    }

    $script:ConsoleLogEnabled = $true

    # ===================== Summary =====================
    Write-Log "=========================================="
    Write-Log "Migration completed"
    Write-Log "Total:        $total"
    Write-Log "Successful:   $okCount"
    Write-Log "Skipped:      $skipCount"
    Write-Log "Errors:       $errCount"
    Write-Log "=========================================="

    if ($errList.Count -gt 0) {
        Write-Log "Error list:" -Level "ERROR"
        foreach ($e in $errList) {
            Write-Log "  $e" -Level "ERROR"
        }
    }

    Write-Log "LogFile: $LogFile"
} finally {
    $script:ConsoleLogEnabled = $true
    $ProgressPreference = $script:OriginalProgressPreference
    Write-Information "" -InformationAction Continue
    [void](Read-Input "Press Enter to close this script...")
}
