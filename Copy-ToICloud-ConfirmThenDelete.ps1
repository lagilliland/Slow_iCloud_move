<#
Copy-ToICloud-ConfirmThenDelete.ps1

- Copy one file at a time from SourceRoot to DestRoot (iCloud Drive folder)
- Poll Explorer "Availability status" until DONE status is observed (stable N polls)
- Then delete source file
- Prune empty source directories (emoji-safe, bottom-up, proven pipeline)
- Progress bar + in-place spinner/dots (no scrolling)
- Timestamped log file
- MaxFiles supports number OR "all" / "*"
- Graceful exit: press ESC to finish current file, then exit
- Polling "elapsed" displays TOTAL RUN TIME for the script
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [Parameter(Mandatory)][string]$SourceRoot,
  [Parameter(Mandatory)][string]$DestRoot,

  # "10", "all", "*"
  [Parameter()][string]$MaxFiles = "10",

  [Parameter()][ValidateRange(1,60)][int]$PollSeconds = 2,
  [Parameter()][ValidateRange(30,86400)][int]$TimeoutSeconds = 900,
  [Parameter()][ValidateRange(1,50)][int]$StablePollsRequired = 2,

  # Treat these as "upload/sync complete" statuses (regex)
  [Parameter()][string]$DoneStatusRegex = '^(Always available on this device|Available on this device)$',

  [Parameter()][string]$LogPath,
  [Parameter()][switch]$RemoveEmptySourceDirs,
  [Parameter()][switch]$PruneWholeSourceTree
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# -------------------------
# Total run stopwatch
# -------------------------
$script:RunSw = [System.Diagnostics.Stopwatch]::StartNew()

# -------------------------
# Graceful ESC exit
# -------------------------
$script:ExitRequested = $false

function Write-Log {
  param(
    [Parameter(Mandatory)][string]$Message,
    [ValidateSet("INFO","WARN","ERROR","POLL")]
    [string]$Level = "INFO"
  )
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
  $line = "[$ts][$Level] $Message"
  Add-Content -LiteralPath $LogPath -Value $line
  if ($Level -in @("WARN","ERROR")) { Write-Host $line }
}

function Test-ExitKeyPressed {
  if ([Console]::KeyAvailable) {
    if ([Console]::ReadKey($true).Key -eq [ConsoleKey]::Escape) {
      if (-not $script:ExitRequested) {
        Write-Log -Level "WARN" "Exit requested by user (ESC). Will finish current file, then exit."
      }
      $script:ExitRequested = $true
    }
  }
}

# -------------------------
# Logging (timestamped)
# -------------------------
if (-not $LogPath) {
  $LogPath = Join-Path $PWD ("icloud-copy-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
} elseif (Test-Path -LiteralPath $LogPath -PathType Container) {
  $LogPath = Join-Path $LogPath ("icloud-copy-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}

# -------------------------
# Spinner (in-place)
# -------------------------
$script:LastLen = 0
function Write-InPlace {
  param([string]$Text)
  $pad = [Math]::Max(0, $script:LastLen - $Text.Length)
  $script:LastLen = $Text.Length
  Write-Host -NoNewline ("`r$Text" + (" " * $pad))
}
function Clear-InPlace {
  Write-InPlace ""
  Write-Host ""
  $script:LastLen = 0
}

function Ensure-Dir {
  param([Parameter(Mandatory)][string]$DirPath)
  if (-not (Test-Path -LiteralPath $DirPath)) {
    New-Item -ItemType Directory -Path $DirPath | Out-Null
  }
}

# -------------------------
# Paths + MaxFiles
# -------------------------
$srcFull = (Resolve-Path -LiteralPath $SourceRoot).Path.TrimEnd('\')
$dstFull = (Resolve-Path -LiteralPath $DestRoot).Path.TrimEnd('\')

$processAll = $false
[int]$maxCount = 0
if ($MaxFiles -match '^(all|\*)$') {
  $processAll = $true
} elseif (-not [int]::TryParse($MaxFiles, [ref]$maxCount) -or $maxCount -lt 1) {
  throw "MaxFiles must be a positive integer or 'all'"
}

# -------------------------
# Explorer Availability Status (Shell.Application)
# -------------------------
$script:Shell = New-Object -ComObject Shell.Application
$script:AvailIdx = @{}

function Get-AvailabilityStatus {
  param([Parameter(Mandatory)][string]$Path)

  $dir  = Split-Path -Parent $Path
  $name = Split-Path -Leaf  $Path
  $sf = $script:Shell.Namespace($dir)
  if (-not $sf) { return "" }

  if (-not $script:AvailIdx.ContainsKey($dir)) {
    $found = $false
    for ($i=0; $i -lt 500; $i++) {
      $hdr = $sf.GetDetailsOf($null,$i)
      if ($hdr -eq "Availability Status" -or $hdr -eq "Availability status") {
        $script:AvailIdx[$dir] = $i; $found = $true; break
      }
      if (-not $found -and $hdr -match 'Availability') {
        $script:AvailIdx[$dir] = $i; $found = $true
      }
    }
    if (-not $script:AvailIdx.ContainsKey($dir)) { $script:AvailIdx[$dir] = 303 }
  }

  $item = $sf.ParseName($name)
  if (-not $item) { return "" }

  return ("$($sf.GetDetailsOf($item, $script:AvailIdx[$dir]))").Trim()
}

# -------------------------
# Wait for DONE status (stable N polls)
# - elapsed shown is TOTAL script runtime
# -------------------------
function Wait-UntilDoneStatus {
  param(
    [Parameter(Mandatory)][string]$DestPath,
    [Parameter(Mandatory)][int]$ParentProgressId
  )

  $inProgressRegex = '^(Sync pending|Syncing|Uploading|Downloading|Pending)$'

  $spinner = @('|','/','-','\')
  $spin = 0
  $dotCount = 0
  $stable = 0

  $fileSw = [System.Diagnostics.Stopwatch]::StartNew()

  while ($true) {
    Test-ExitKeyPressed

    $status = Get-AvailabilityStatus -Path $DestPath

    $isBlank = [string]::IsNullOrWhiteSpace($status)
    $isInProgress = (-not $isBlank) -and ($status -match $inProgressRegex)
    $isDone = (-not $isBlank) -and ($status -match $DoneStatusRegex)

    if ($isDone) { $stable++ } else { $stable = 0 }

    $dots = "." * ($dotCount % 11); $dotCount++
    $sp = $spinner[$spin++ % 4]

    $runElapsed  = [TimeSpan]::FromSeconds([int]$script:RunSw.Elapsed.TotalSeconds).ToString("hh\:mm\:ss")
    $fileElapsed = [TimeSpan]::FromSeconds([int]$fileSw.Elapsed.TotalSeconds).ToString("hh\:mm\:ss")

    $displayStatus = if ($isBlank) { "(blank)" } else { $status }

    # In-place overlay
    Write-InPlace ("Polling {0} {1} {2} | done-stable {3}/{4} | run {5} | file {6}" -f `
      $sp, ($displayStatus -replace "`r|`n"," "), $dots, $stable, $StablePollsRequired, $runElapsed, $fileElapsed)

    # Sub progress
    Write-Progress -Id 2 -ParentId $ParentProgressId `
      -Activity "Waiting for iCloud sync" `
      -Status ("Status: {0} | run {1} | file {2}" -f $displayStatus, $runElapsed, $fileElapsed) `
      -PercentComplete 0

    Write-Log -Level "POLL" -Message ("Availability='{0}' Blank={1} InProgress={2} Done={3} DoneStable={4}/{5} Run={6} File={7} Path={8}" -f `
      $status, $isBlank, $isInProgress, $isDone, $stable, $StablePollsRequired, $runElapsed, $fileElapsed, $DestPath)

    if ($stable -ge $StablePollsRequired) {
      Clear-InPlace
      Write-Progress -Id 2 -ParentId $ParentProgressId -Activity "Waiting for iCloud sync" -Completed
      return $true
    }

    if ($fileSw.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
      Clear-InPlace
      Write-Progress -Id 2 -ParentId $ParentProgressId -Activity "Waiting for iCloud sync" -Completed
      return $false
    }

    Start-Sleep -Seconds $PollSeconds
  }
}

# -------------------------
# Empty-dir prune (emoji-safe, bottom-up, proven pipeline)
# - plus: attempt scope dir, then walk parents up to root
# -------------------------
function Prune-EmptyDirs {
  param(
    [Parameter(Mandatory)][string]$RootPath,
    [Parameter(Mandatory)][string]$ScopePath
  )

  $root = (Resolve-Path -LiteralPath $RootPath).Path.TrimEnd('\')

  if (-not (Test-Path -LiteralPath $ScopePath -PathType Container)) {
    # Scope might already be gone; nothing to do
    Write-Log -Level "POLL" -Message ("Prune: scope missing, skipping. Scope={0}" -f $ScopePath)
    return
  }

  # 1) Delete empty dirs under scope (deepest first) — this is your working pattern
  try {
    Get-ChildItem -LiteralPath $ScopePath -Directory -Recurse -Force -ErrorAction SilentlyContinue |
      Sort-Object FullName -Descending |
      Where-Object { -not (Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue | Select-Object -First 1) } |
      ForEach-Object {
        try {
          if ($PSCmdlet.ShouldProcess($_.FullName, "Remove empty directory")) {
            Write-Log -Message ("Removing empty directory: {0}" -f $_.FullName)
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
          }
        } catch {
          Write-Log -Level "WARN" -Message ("Failed to remove dir: {0} | {1}" -f $_.FullName, $_.Exception.Message)
        }
      }
  } catch {
    Write-Log -Level "WARN" -Message ("Prune enumeration failed for scope: {0} | {1}" -f $ScopePath, $_.Exception.Message)
  }

  # 2) Attempt to delete the scope itself if empty
  try {
    if ((Test-Path -LiteralPath $ScopePath -PathType Container) -and -not (Get-ChildItem -LiteralPath $ScopePath -Force -ErrorAction SilentlyContinue | Select-Object -First 1)) {
      if ($PSCmdlet.ShouldProcess($ScopePath, "Remove empty directory")) {
        Write-Log -Message ("Removing empty directory: {0}" -f $ScopePath)
        Remove-Item -LiteralPath $ScopePath -Force -ErrorAction Stop
      }
    }
  } catch {
    Write-Log -Level "WARN" -Message ("Failed to remove scope dir: {0} | {1}" -f $ScopePath, $_.Exception.Message)
  }

  # 3) Walk upward and remove empty parents up to (but NOT including) RootPath
  $current = Split-Path -Parent $ScopePath
  while ($current -and ($current.TrimEnd('\') -ne $root)) {
    try {
      if (-not (Test-Path -LiteralPath $current -PathType Container)) { break }

      $hasAnything = Get-ChildItem -LiteralPath $current -Force -ErrorAction SilentlyContinue | Select-Object -First 1
      if ($hasAnything) { break }

      if ($PSCmdlet.ShouldProcess($current, "Remove empty directory")) {
        Write-Log -Message ("Removing empty directory: {0}" -f $current)
        Remove-Item -LiteralPath $current -Force -ErrorAction Stop
      }
    } catch {
      Write-Log -Level "WARN" -Message ("Failed to remove parent dir: {0} | {1}" -f $current, $_.Exception.Message)
      break
    }

    $current = Split-Path -Parent $current
  }
}

# -------------------------
# Start
# -------------------------
Write-Host "Press ESC to finish the current file and exit." -ForegroundColor DarkGray
Write-Log "Run started."
Write-Log ("SourceRoot={0}" -f $srcFull)
Write-Log ("DestRoot={0}" -f $dstFull)
Write-Log ("MaxFiles={0}" -f $MaxFiles)
Write-Log ("PollSeconds={0} TimeoutSeconds={1} StablePollsRequired={2}" -f $PollSeconds, $TimeoutSeconds, $StablePollsRequired)
Write-Log ("DoneStatusRegex={0}" -f $DoneStatusRegex)
Write-Log ("RemoveEmptySourceDirs={0} PruneWholeSourceTree={1}" -f $RemoveEmptySourceDirs, $PruneWholeSourceTree)
Write-Log ("LogPath={0}" -f $LogPath)

$all = Get-ChildItem -LiteralPath $srcFull -File -Recurse | Sort-Object FullName
$files = if ($processAll) { $all } else { $all | Select-Object -First $maxCount }

$progressId = 1
$i = 0

foreach ($f in $files) {

  if ($script:ExitRequested) {
    Write-Log -Level "WARN" "Exit requested. No new files will be started."
    break
  }

  $i++
  $rel = $f.FullName.Substring($srcFull.Length).TrimStart('\')
  $dest = Join-Path $dstFull $rel
  $destDir = Split-Path -Parent $dest

  $pct = if ($files.Count -gt 0) { [int](($i / [double]$files.Count) * 100) } else { 0 }
  $runElapsed = [TimeSpan]::FromSeconds([int]$script:RunSw.Elapsed.TotalSeconds).ToString("hh\:mm\:ss")

  Write-Progress -Id $progressId `
    -Activity "Copying to iCloud" `
    -Status ("{0}/{1}: {2} | run {3}" -f $i, $files.Count, $rel, $runElapsed) `
    -PercentComplete $pct

  try {
    Ensure-Dir -DirPath $destDir

    Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
    Write-Log ("Copied: {0} -> {1}" -f $f.FullName, $dest)

    $ok = Wait-UntilDoneStatus -DestPath $dest -ParentProgressId $progressId

    if ($ok) {
      Remove-Item -LiteralPath $f.FullName -Force
      Write-Log "Deleted source."

      if ($RemoveEmptySourceDirs) {
        $scope = if ($PruneWholeSourceTree) { $srcFull } else { (Split-Path -Parent $f.FullName) }
        Write-Log -Level "POLL" -Message ("Prune scope={0}" -f $scope)
        Prune-EmptyDirs -RootPath $srcFull -ScopePath $scope
      }
    } else {
      Write-Log -Level "WARN" "Timeout waiting for DONE status. Source preserved."
    }
  }
  catch {
    Write-Log -Level "ERROR" ("Failed: {0} | {1}" -f $f.FullName, $_.Exception.Message)
    continue
  }
}

Write-Progress -Id $progressId -Activity "Copying to iCloud" -Completed
Write-Progress -Id 2 -Activity "Waiting for iCloud sync" -Completed

$finalRun = [TimeSpan]::FromSeconds([int]$script:RunSw.Elapsed.TotalSeconds).ToString("hh\:mm\:ss")
Write-Log ("Run complete. TotalRunTime={0}" -f $finalRun)
