[CmdletBinding()]
param(
  [string]$source,
  [string]$dest,
  [Alias('Path7z', 'zpath')]
  [string]$SevenZipPath,
  [switch]$yes,
  [switch]$no,
  [switch]$nolog,
  [switch]$ignore,
  [switch]$help
)

# Require PowerShell 7.6.0+ (before StrictMode so older hosts get a clear message)
$minPowerShell = [version]'7.6.0'
if ($PSVersionTable.PSVersion -lt $minPowerShell) {
  Write-Host "The minimal requirement is PowerShell 7.6.0 (current: $($PSVersionTable.PSVersion)). Install pwsh from GitHub releases." -ForegroundColor Red
  Start-Process 'https://github.com/powershell/powershell/releases'
  exit 1
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:lastStatusLineLength = 0

function Write-Fail {
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Message
  )
  if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = 'An unexpected error occurred.'
  }
  Write-Host $Message -ForegroundColor Red
  exit 1
}

function Clear-HostSafe {
  try { Clear-Host } catch { }
}

function Show-Header {
  Write-Host "=== [ CHD-Maid ] ===" -ForegroundColor Green
  Write-Host "=== [ Version 2026.4.3 ] ===" -ForegroundColor Yellow
  Write-Host ""
}

function Write-Info {
  param([Parameter(Mandatory = $true)][string]$Message)
  Write-Host $Message -ForegroundColor White
}

function Write-Warn {
  param([Parameter(Mandatory = $true)][string]$Message)
  Write-Host $Message -ForegroundColor Yellow
}

function Write-Summary {
  param([Parameter(Mandatory = $true)][string]$Message)
  Write-Host $Message -ForegroundColor DarkCyan
}

function Write-MessageWithFlags {
  param(
    [Parameter(Mandatory = $true)][string]$Text,
    [ConsoleColor]$Color = [ConsoleColor]::White,
    [ConsoleColor]$FlagColor = [ConsoleColor]::Cyan,
    [ConsoleColor]$ParenColor = [ConsoleColor]::Yellow
  )
  $parts = [regex]::Split($Text, '(\s-[A-Za-z0-9]+|\([^)]*\))')
  foreach ($part in $parts) {
    if ($part -eq '') { continue }
    if ($part -match '^(\s+)(-[A-Za-z0-9]+)$') {
      Write-Host $matches[1] -NoNewline -ForegroundColor $Color
      Write-Host $matches[2] -NoNewline -ForegroundColor $FlagColor
    } elseif ($part -match '^\([^)]*\)$') {
      Write-Host $part -NoNewline -ForegroundColor $ParenColor
    } else {
      Write-Host $part -NoNewline -ForegroundColor $Color
    }
  }
  Write-Host ""
}

function Format-TextPreview {
  param(
    [Parameter(Mandatory = $true)][string]$Text,
    [Parameter(Mandatory = $true)][int]$MaxLength
  )
  if ($MaxLength -le 0) { return '' }
  if ($Text.Length -le $MaxLength) { return $Text }
  if ($MaxLength -le 3) { return $Text.Substring(0, $MaxLength) }
  return $Text.Substring(0, $MaxLength - 3) + '...'
}

function Get-DoubleQuotedPath {
  param([string]$Path)
  if ([string]::IsNullOrEmpty($Path)) { return '""' }
  $t = $Path.Trim()
  if ($t.Length -ge 2 -and $t.StartsWith('"') -and $t.EndsWith('"')) {
    return $t
  }
  $escaped = $Path.Replace('"', '""')
  return '"' + $escaped + '"'
}

# Completion summary: e.g. 0.54 GB (556.39 MB); 1024-based (GiB/MiB) for consistency with Windows binary prefixes.
# Force en-US so thousands separators always match US rules (e.g. 4,200.50 MB).
function Format-DataSizeGbMb {
  param([long]$Bytes)
  $negative = $Bytes -lt 0
  $abs = if ($negative) { [long][math]::Abs([double]$Bytes) } else { $Bytes }
  $gb = [double]$abs / [math]::Pow(1024, 3)
  $mb = [double]$abs / [math]::Pow(1024, 2)
  $c = [cultureinfo]::GetCultureInfo('en-US')
  $out = [string]::Format($c, '{0:N2} GB ({1:N2} MB)', $gb, $mb)
  if ($negative) { return '-' + $out }
  return $out
}

function Format-UsInt32 {
  param([int]$Value)
  return $Value.ToString('N0', [cultureinfo]::GetCultureInfo('en-US'))
}

function Add-ChdmaidPerFileLogLine {
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Line
  )
  if ($null -ne $script:chdmaidPerFileLogLines) {
    [void]$script:chdmaidPerFileLogLines.Add($Line)
  }
}

function Complete-ChdmaidPerFileLog {
  param([Parameter(Mandatory = $true)][bool]$Persist)
  if (-not $Persist) { return }
  if ($null -eq $script:chdmaidPerFileLogLines -or $script:chdmaidPerFileLogLines.Count -eq 0) { return }
  Add-Content -LiteralPath $logPath -Encoding utf8 -Value $script:chdmaidPerFileLogLines.ToArray()
  Add-Content -LiteralPath $logPath -Encoding utf8 -Value ''
}

# Remove the unpack tree for this archive pass before any sibling source item runs (finally in Expand-OneArchiveAndVisit). The archive file on disk is removed only with -yes ($deleteSources). Nested archives must be gone; with -yes, no disc images may remain (sources were deleted); with -no, disc files may still exist until this deletes the whole tree.
function Remove-ArchiveExtractionIfReady {
  param(
    [Parameter(Mandatory = $true)][string]$ExtractDir,
    [AllowNull()]
    [AllowEmptyString()]
    [string]$ArchivePath,
    [Parameter(Mandatory = $true)][hashtable]$InputOutcome,
    [Parameter(Mandatory = $true)][string[]]$AllInputPaths
  )
  if ([string]::IsNullOrWhiteSpace($ExtractDir)) { return }
  try {
    $dirNorm = [IO.Path]::GetFullPath($ExtractDir.TrimEnd('\').TrimEnd('/'))
  } catch {
    $dirNorm = $ExtractDir.TrimEnd('\').TrimEnd('/')
  }
  $sep = [IO.Path]::DirectorySeparatorChar
  $prefix = $dirNorm + $sep
  # Use a list so a single path is not iterated as characters; FullName must match InputOutcome keys.
  $related = [System.Collections.Generic.List[string]]::new()
  foreach ($ip in $AllInputPaths) {
    if ([string]::IsNullOrWhiteSpace($ip)) { continue }
    try {
      $ipNorm = [IO.Path]::GetFullPath($ip)
    } catch {
      $ipNorm = $ip
    }
    if (
      $ipNorm.Equals($dirNorm, [StringComparison]::OrdinalIgnoreCase) -or
      $ipNorm.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)
    ) {
      [void]$related.Add($ip)
    }
  }
  foreach ($rf in $related) {
    if ($InputOutcome[$rf] -eq 'failed') {
      return
    }
  }
  if (-not (Test-Path -LiteralPath $dirNorm -PathType Container)) {
    if ($deleteSources -and -not [string]::IsNullOrWhiteSpace($ArchivePath) -and (Test-Path -LiteralPath $ArchivePath -PathType Leaf)) {
      Write-Host "Removing archive: " -NoNewline -ForegroundColor DarkCyan
      Write-Host $ArchivePath -ForegroundColor DarkGray
      try {
        Remove-Item -LiteralPath $ArchivePath -Force -ErrorAction Stop
      } catch {
        Write-Warn ("Could not remove archive: {0} | {1}" -f $ArchivePath, $_.Exception.Message)
      }
    }
    return
  }
  $foundArchiveLeft = $false
  $foundDiscLeft = $false
  foreach ($leaf in Get-ChildItem -LiteralPath $dirNorm -Recurse -File -ErrorAction SilentlyContinue) {
    $nm = $leaf.Name
    if (Test-IsArchiveFileName -FileName $nm) {
      $foundArchiveLeft = $true
      break
    }
    if ($deleteSources -and (Test-IsDiscInputFileName -FileName $nm)) {
      $foundDiscLeft = $true
    }
  }
  if ($foundArchiveLeft) { return }
  if ($deleteSources -and $foundDiscLeft) { return }
  Write-Host "Removing extract folder: " -NoNewline -ForegroundColor DarkCyan
  Write-Host $dirNorm -ForegroundColor DarkGray
  try {
    Remove-Item -LiteralPath $dirNorm -Recurse -Force -ErrorAction Stop
  } catch {
    Write-Warn ("Could not remove extract folder: {0} | {1}" -f $dirNorm, $_.Exception.Message)
  }
  if ($deleteSources -and -not [string]::IsNullOrWhiteSpace($ArchivePath) -and (Test-Path -LiteralPath $ArchivePath -PathType Leaf)) {
    Write-Host "Removing archive: " -NoNewline -ForegroundColor DarkCyan
    Write-Host $ArchivePath -ForegroundColor DarkGray
    try {
      Remove-Item -LiteralPath $ArchivePath -Force -ErrorAction Stop
    } catch {
      Write-Warn ("Could not remove archive: {0} | {1}" -f $ArchivePath, $_.Exception.Message)
    }
  }
}

function Remove-CompletedArchiveExtractions {
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [object[]]$ExtractRecords,
    [Parameter(Mandatory = $true)][hashtable]$InputOutcome,
    [Parameter(Mandatory = $true)][string[]]$AllInputPaths
  )
  if ($null -eq $ExtractRecords -or @($ExtractRecords).Count -eq 0) { return }
  $sorted = @($ExtractRecords | Sort-Object { $_.ExtractDir.Length } -Descending)
  foreach ($rec in $sorted) {
    Remove-ArchiveExtractionIfReady -ExtractDir $rec.ExtractDir -ArchivePath $rec.ArchivePath -InputOutcome $InputOutcome -AllInputPaths $AllInputPaths
  }
}

function Write-StatusLine {
  param(
    [Parameter(Mandatory = $true)][string]$Label,
    [Parameter(Mandatory = $true)][string]$FileName,
    [Parameter(Mandatory = $true)][TimeSpan]$Elapsed,
    [Parameter(Mandatory = $false)][int]$Percent = 0
  )
  $elapsedText = $Elapsed.ToString('hh\:mm\:ss')
  $pct = [Math]::Min(100, [Math]::Max(0, $Percent))
  $pctText = '{0:D3}' -f $pct
  $progressPrefix = "[$elapsedText] ${pctText}% - ${Label} - "

  $consoleWidth = 120
  try { $consoleWidth = $Host.UI.RawUI.WindowSize.Width } catch { }
  $maxFileNameLength = [Math]::Max(8, $consoleWidth - $progressPrefix.Length)
  $fileNameDisplay = Format-TextPreview -Text $FileName -MaxLength $maxFileNameLength

  $line = $progressPrefix + $fileNameDisplay
  $pad = ' ' * [Math]::Max(0, $script:lastStatusLineLength - $line.Length)

  Write-Host "`r$progressPrefix" -NoNewline -ForegroundColor White
  Write-Host $fileNameDisplay -NoNewline -ForegroundColor DarkGray
  Write-Host $pad -NoNewline -ForegroundColor DarkGray

  $script:lastStatusLineLength = ($line.Length)
}

function Show-Help {
  Clear-HostSafe
  Show-Header
  
  Write-Host "Summary:" -ForegroundColor DarkYellow
  Write-Host 'Creates CHD images from .cue, .gdi, and .iso using chdman (createcd + verify).' -ForegroundColor White
  Write-Host "Scans -source (including folders created there when archives are extracted)." -ForegroundColor White
  Write-Host ""

  Write-Host "Requirements:" -ForegroundColor DarkYellow
  Write-Info "  - PowerShell 7.6.0 or newer."
  Write-Info "  - chdman.exe must be in the current directory"
  Write-Info "    -- This is included in releases on the official GitHub page"
  Write-Info "  - 7-Zip is required for handling archives."
  Write-Info "    -- 7-Zip 26.00 (2026-02-12) used during development."
  Write-Host ""

  Write-Host "Archives (`.zip` / `.7z` / `.rar`):" -ForegroundColor DarkYellow
  Write-MessageWithFlags "  - Matched files under `-source` extract into `-source`\<archive name>\ (not under `-dest`)."
  Write-Info "    -- Each archive is extracted in source-tree order, then converted and verified before the next source item."
  Write-MessageWithFlags "  - If that folder name already exists under `-source`, `<name> (2)`, `(3)`, ... is used."
  Write-Info "    -- When disc jobs under an archive's unpack folder all finish without failure, that extracted folder is removed before the next source item."
  Write-MessageWithFlags "    -- `-yes` also deletes the archive file; `-no` keeps the archive file."
  Write-MessageWithFlags "    -- If neither `-no` or `-yes` are passed, `-no` is assumed."
  Write-MessageWithFlags "  - Output `.chd` files are always written to the top level of `-dest` (e.g. `-dest\ImageName.chd`)."
  Write-Host ""

  Write-Host "Source & Destination:" -ForegroundColor DarkYellow
  Write-MessageWithFlags "  `-source` must be an existing folder." 
  Write-Host "   -- If you pass a bad path, you are reprompted until it is valid."
  Write-MessageWithFlags "  `-dest` is created if it does not exist."
  Write-Host "   -- You can type a new folder at the prompt."
  Write-Info "  - Pressing <Enter> uses the current directory."
  Write-Info "  - If you omit both `-yes` and `-ignore`, you are prompted whether existing destination `.chd` files should be [I]gnored or [R]escanned."
  Write-Info "   -- Enter chooses Rescan (verify / replace if bad)."
  Write-Host ""

  Write-Host "Logging:" -ForegroundColor DarkYellow
  Write-Info "  - Default log: `logs\log-chd-maid-YYYY-MM-dd-HHmmss.log` under the current directory (UTF-8 formatted)."
  Write-Info "    -- `logs` subdirectory is created if missing."
  Write-Info "  - The completion summary is appended to the log when logging is enabled or any job failed."
  Write-Info "    -- The console completion summary ends with the log file path when a log exists (easier to find after long runs)."
  Write-Info "    -- While logging is enabled (or on failure), each disc job is written as one block of lines after that job finishes."
  Write-MessageWithFlags "  `-nolog` no log file is generated for a fully successful run."
  Write-Info "   -- On the first conversion failure, a log file is created with a short header with FAIL details. "
  Write-Host ""

  Write-Host "Parameters:" -ForegroundColor DarkYellow
  Write-MessageWithFlags "  -source <path>        Root folder to scan recursively for archives and for `.cue` / `.gdi` / `.iso`."
  Write-MessageWithFlags "  -dest <path>          Output root for `.chd` only (archives unpack under `-source`)."
  Write-MessageWithFlags "  -SevenZipPath <path>  Optional `7z.exe` or its parent folder. Aliases: `-Path7z`, `-zpath`."
  Write-MessageWithFlags "  `-ignore`               Skip when the matching `.chd` already exists under `-dest` (no verify)."
  Write-MessageWithFlags "   -- For archives, using the archive base name (e.g. `ImageName.7z` expects `ImageName.chd`)."
  Write-MessageWithFlags "   -- Same as choosing [I] at the Ignore/Rescan prompt when you do not pass `-yes`."
  Write-MessageWithFlags "   -- Using `-ignore` implies keeping sources: not allowed with `-yes`."
  Write-Host ""

  Write-Host "Examples:" -ForegroundColor DarkYellow
  Write-MessageWithFlags "  .\chd-maid.ps1 -source `"D:\input`" -dest `"D:\output`" -Path7z `"D:\7-Zip`""
  Write-MessageWithFlags "  .\chd-maid.ps1 -source `"D:\input`" -dest `"D:\output`" -yes"
  Write-MessageWithFlags "  .\chd-maid.ps1 -source `"D:\input`" -dest `"D:\output`" -no"
  Write-MessageWithFlags "  .\chd-maid.ps1 -source `"D:\input`" -dest `"D:\output`" -no -nolog"
  Write-MessageWithFlags "  .\chd-maid.ps1 -source `"D:\input`" -dest `"D:\output`" -no -ignore"
  Write-Host ""
}

# PS 7.6+ with StrictMode: `$args` may be absent when `-File` is used with no unbound arguments.
$argsVar = Get-Variable -Name args -Scope Script -ErrorAction SilentlyContinue
$remainingArgs = if ($argsVar) { @($argsVar.Value) } else { @() }
$helpTokens = @(
  foreach ($a in $remainingArgs) {
    if ($null -eq $a) { continue }
    [string]$t = $a
    # Normalize values like "'-help'" or "\"-help\"" coming through wrappers/ArgumentList
    $t.Trim('''','"')
  }
)

if (
  $help -or
  ($helpTokens -contains '-help') -or
  ($helpTokens -contains '--help') -or
  ($helpTokens -contains '-?') -or
  ($helpTokens -contains '/help')
) {
  Show-Help
  return
}

Clear-HostSafe
Show-Header

if ($yes -and $no) {
  throw "Use can only use one -yes (delete source) or -no (do not delete source)."
}

if ($ignore -and $yes) {
  Write-Fail '-ignore cannot be used with -yes. Use -ignore with -no to keep all original sources and archives intact.'
}

$writeFullLog = -not $nolog

$deleteSources = $false
if ($yes) { $deleteSources = $true }

# Source must exist as a directory. If -source is invalid, reprompt until valid.
function Resolve-SourceFolderParameter {
  param(
    [Parameter(Mandatory = $true)][string]$PromptLabel,
    [string]$BoundValue,
    [Parameter(Mandatory = $true)][bool]$WasBound
  )
  $defaultPath = (Get-Location).Path
  if ($WasBound) {
    $folder = if ([string]::IsNullOrWhiteSpace($BoundValue)) { $defaultPath } else { $BoundValue.Trim() }
  } else {
    $folder = $null
  }
  while ($true) {
    if (-not [string]::IsNullOrWhiteSpace($folder) -and (Test-Path -LiteralPath $folder -PathType Container)) {
      return $folder
    }
    if (-not [string]::IsNullOrWhiteSpace($folder)) {
      Write-Warning "Path not found or not a folder: $folder"
    }
    $folder = Read-Host "$PromptLabel enter to use default (currently: $defaultPath)"
    if ([string]::IsNullOrWhiteSpace($folder)) {
      $folder = $defaultPath
    } else {
      $folder = $folder.Trim()
    }
  }
}

# Destination may not exist yet; it is created later. Interactive prompt does not require a pre-existing folder.
function Resolve-DestFolderParameter {
  param(
    [Parameter(Mandatory = $true)][string]$PromptLabel,
    [string]$BoundValue,
    [Parameter(Mandatory = $true)][bool]$WasBound
  )
  $defaultPath = (Get-Location).Path
  if ($WasBound) {
    if ([string]::IsNullOrWhiteSpace($BoundValue)) {
      return $defaultPath
    }
    return $BoundValue.Trim()
  }
  $folder = Read-Host "$PromptLabel enter to use default (currently: $defaultPath)"
  if ([string]::IsNullOrWhiteSpace($folder)) {
    return $defaultPath
  }
  return $folder.Trim()
}

function Resolve-SevenZipExecutable {
  param(
    [string]$BoundValue,
    [Parameter(Mandatory = $true)][bool]$WasBound,
    [Parameter(Mandatory = $true)][string]$ExecutedDir
  )
  $sevenZipOfficialDownload = 'https://www.7-zip.org/download.html'
  $programFiles7z = Join-Path 'C:\Program Files\7-Zip' '7z.exe'
  if (-not $WasBound -or [string]::IsNullOrWhiteSpace($BoundValue)) {
    $local7z = Join-Path $ExecutedDir '7z.exe'
    if (Test-Path -LiteralPath $local7z -PathType Leaf) {
      return [IO.Path]::GetFullPath($local7z)
    }
    if (Test-Path -LiteralPath $programFiles7z -PathType Leaf) {
      return [IO.Path]::GetFullPath($programFiles7z)
    }
    Write-Host "7-Zip is required to extract archives. Opening the official download page..." -ForegroundColor Yellow
    Start-Process $sevenZipOfficialDownload
    Write-Fail "7-Zip (7z.exe) is required when the source folder contains archives. It was not found in the current directory ($ExecutedDir) or at $programFiles7z. After installing, copy 7z.exe next to chd-maid/chdman if you prefer, or use -SevenZipPath / -Path7z."
  }
  $trimmed = $BoundValue.Trim().Trim('"')
  $p = $trimmed
  if ((Test-Path -LiteralPath $p -PathType Container)) {
    $p = Join-Path $p '7z.exe'
  } elseif (-not (Test-Path -LiteralPath $p -PathType Leaf)) {
    $tryExe = Join-Path $trimmed '7z.exe'
    if (Test-Path -LiteralPath $tryExe -PathType Leaf) {
      $p = $tryExe
    }
  }
  if (-not (Test-Path -LiteralPath $p -PathType Leaf)) {
    Write-Fail "7-Zip executable not found (expected 7z.exe): $BoundValue"
  }
  return [IO.Path]::GetFullPath($p)
}

$source = Resolve-SourceFolderParameter -PromptLabel 'Source folder' -BoundValue $source -WasBound $PSBoundParameters.ContainsKey('source')
$dest = Resolve-DestFolderParameter -PromptLabel 'Destination folder' -BoundValue $dest -WasBound $PSBoundParameters.ContainsKey('dest')

Write-Host ""

# Locate chdman.exe
$executedDir = (Get-Location).Path
$chdmanExe = Join-Path $executedDir "chdman.exe"
if (-not (Test-Path -LiteralPath $chdmanExe -PathType Leaf)) {
  Write-Fail "chdman.exe does not exist in the executed directory: $executedDir"
}

New-Item -ItemType Directory -Path $dest -Force | Out-Null

function Test-IsArchiveFileName {
  param([Parameter(Mandatory = $true)][string]$FileName)
  $n = $FileName.ToLowerInvariant()
  return ($n.EndsWith('.zip') -or $n.EndsWith('.7z') -or $n.EndsWith('.rar'))
}

function Test-IsDiscInputFileName {
  param([Parameter(Mandatory = $true)][string]$FileName)
  $n = $FileName.ToLowerInvariant()
  return ($n.EndsWith('.cue') -or $n.EndsWith('.gdi') -or $n.EndsWith('.iso'))
}

function Get-ArchiveStemForFolderName {
  param([Parameter(Mandatory = $true)][string]$FileName)
  return [IO.Path]::GetFileNameWithoutExtension($FileName)
}

function Get-UniqueExtractDirectoryPath {
  param(
    [Parameter(Mandatory = $true)][string]$ParentRoot,
    [Parameter(Mandatory = $true)][string]$FolderStem
  )
  $candidate = Join-Path $ParentRoot $FolderStem
  if (-not (Test-Path -LiteralPath $candidate)) {
    return [IO.Path]::GetFullPath($candidate)
  }
  $i = 2
  while ($true) {
    $alt = Join-Path $ParentRoot ('{0} ({1})' -f $FolderStem, $i)
    if (-not (Test-Path -LiteralPath $alt)) {
      return [IO.Path]::GetFullPath($alt)
    }
    $i++
  }
}

function Invoke-ProcessWithStatus {
  param(
    [Parameter(Mandatory = $true)][string]$ExePath,
    [Parameter(Mandatory = $true)][string[]]$Arguments,
    [Parameter(Mandatory = $true)][string]$Label,
    [Parameter(Mandatory = $true)][string]$FileName,
    [Parameter(Mandatory = $true)][System.Diagnostics.Stopwatch]$Stopwatch
  )

  $psi = [System.Diagnostics.ProcessStartInfo]::new()
  $psi.FileName = $ExePath
  foreach ($a in $Arguments) { [void]$psi.ArgumentList.Add($a) }
  $psi.UseShellExecute = $false
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.CreateNoWindow = $true

  $p = [System.Diagnostics.Process]::new()
  $p.StartInfo = $psi

  $streamState = [hashtable]::Synchronized(@{
    Percent = 0
    StdOut  = [System.Text.StringBuilder]::new()
    StdErr  = [System.Text.StringBuilder]::new()
  })
  # Register-ObjectEvent actions run in a separate runspace; $MessageData is unreliable there—use a global sync state.
  $global:ChdmaidProcessStreamState = $streamState

  try {
    $null = Register-ObjectEvent -InputObject $p -EventName OutputDataReceived -Action {
      $line = $EventArgs.Data
      if ($null -eq $line) { return }
      $st = $global:ChdmaidProcessStreamState
      [void]$st.StdOut.AppendLine($line)
      # chdman prints compression "ratio=100.0%" on the same line as progress; that is not job %.
      $chunk = $line
      $ratioCut = [regex]::Match($line, '(?i)ratio\s*=')
      if ($ratioCut.Success) { $chunk = $line.Substring(0, $ratioCut.Index) }
      if ($chunk -match '(?i)-nan\s*%') { return }
      if ($chunk -match '(\d+(?:\.\d+)?)\s*%') {
        $v = [double]$matches[1]
        if (-not ([double]::IsNaN($v) -or [double]::IsInfinity($v))) {
          $newPct = [int][math]::Round([math]::Min(100, [math]::Max(0, $v)))
          $st.Percent = [math]::Max($st.Percent, $newPct)
        }
      }
    }
    $null = Register-ObjectEvent -InputObject $p -EventName ErrorDataReceived -Action {
      $line = $EventArgs.Data
      if ($null -eq $line) { return }
      $st = $global:ChdmaidProcessStreamState
      [void]$st.StdErr.AppendLine($line)
      $chunk = $line
      $ratioCut = [regex]::Match($line, '(?i)ratio\s*=')
      if ($ratioCut.Success) { $chunk = $line.Substring(0, $ratioCut.Index) }
      if ($chunk -match '(?i)-nan\s*%') { return }
      if ($chunk -match '(\d+(?:\.\d+)?)\s*%') {
        $v = [double]$matches[1]
        if (-not ([double]::IsNaN($v) -or [double]::IsInfinity($v))) {
          $newPct = [int][math]::Round([math]::Min(100, [math]::Max(0, $v)))
          $st.Percent = [math]::Max($st.Percent, $newPct)
        }
      }
    }

    [void]$p.Start()
    $p.BeginOutputReadLine()
    $p.BeginErrorReadLine()

    while (-not $p.HasExited) {
      Write-StatusLine -Label $Label -FileName $FileName -Elapsed $Stopwatch.Elapsed -Percent $streamState.Percent
      Start-Sleep -Milliseconds 250
    }

    $p.WaitForExit()
    Start-Sleep -Milliseconds 150
    # chdman often ends on 99.x% in output; on success show 100% for the final line.
    if ($p.ExitCode -eq 0) {
      $streamState.Percent = 100
    }
    Write-StatusLine -Label $Label -FileName $FileName -Elapsed $Stopwatch.Elapsed -Percent $streamState.Percent

    return [pscustomobject]@{
      ExitCode = $p.ExitCode
      StdOut   = $streamState.StdOut.ToString()
      StdErr   = $streamState.StdErr.ToString()
    }
  } finally {
    Get-EventSubscriber -ErrorAction SilentlyContinue |
      Where-Object { $_.SourceObject -eq $p } |
      ForEach-Object { Unregister-Event -SubscriptionId $_.SubscriptionId -ErrorAction SilentlyContinue }
    $global:ChdmaidProcessStreamState = $null
    $p.Dispose()
  }
}

function Expand-ArchiveWith7Zip {
  param(
    [Parameter(Mandatory = $true)][string]$LiteralArchivePath,
    [Parameter(Mandatory = $true)][string]$ExtractDir,
    [Parameter(Mandatory = $true)][string]$SevenZipExe
  )
  New-Item -ItemType Directory -Path $ExtractDir -Force | Out-Null
  if (-not (Test-Path -LiteralPath $SevenZipExe -PathType Leaf)) {
    throw "7z.exe not found: $SevenZipExe"
  }
  $argOut = '-o' + $ExtractDir + [IO.Path]::DirectorySeparatorChar
  $archiveName = [IO.Path]::GetFileName($LiteralArchivePath)
  $extractStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
  # -bsp1: include percent in stdout so Write-StatusLine can track progress (same pattern as chdman).
  $exResult = Invoke-ProcessWithStatus -ExePath $SevenZipExe -Arguments @('x', $LiteralArchivePath, $argOut, '-y', '-bsp1') -Label 'Extracting' -FileName $archiveName -Stopwatch $extractStopwatch
  Write-Host ""
  if ($exResult.ExitCode -ne 0) {
    throw "7z failed (exit $($exResult.ExitCode)): $LiteralArchivePath"
  }
}

$existingChdsInDest = @()
if (Test-Path -LiteralPath $dest) {
  $existingChdsInDest = @(Get-ChildItem -LiteralPath $dest -Recurse -File -Filter '*.chd' -ErrorAction SilentlyContinue | Sort-Object FullName)
  if ($existingChdsInDest -isnot [System.Array]) {
    $existingChdsInDest = @($existingChdsInDest)
  }
}

function Get-CueReferencedBins {
  param(
    [Parameter(Mandatory = $true)]
    [string]$cuePath
  )

  $cueDir = Split-Path -Parent $cuePath
  $bins = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

  # Parse FILE "name.bin" BINARY lines so we can delete referenced bins when -yes is used.
  # chdman itself will still resolve and join the bins from the cue file.
  foreach ($line in (Get-Content -LiteralPath $cuePath -ErrorAction SilentlyContinue)) {
    $binName = $null

    # FILE "name.bin" BINARY
    if ($line -match '^\s*FILE\s+"([^"]+)"') {
      $binName = $matches[1].Trim()
    }
    # FILE 'name.bin' BINARY (some cues)
    elseif ($line -match "^\s*FILE\s+'([^']+)'") {
      $binName = $matches[1].Trim()
    }
    # FILE name.bin BINARY (unquoted)
    elseif ($line -match '^\s*FILE\s+([^\s]+)\s+BINARY') {
      $binName = $matches[1].Trim().Trim('"').Trim("'")
    }

    if ([string]::IsNullOrWhiteSpace($binName)) { continue }

    if ([IO.Path]::IsPathRooted($binName)) {
      $binPath = $binName
    } else {
      $binPath = Join-Path $cueDir $binName
    }

    if (Test-Path -LiteralPath $binPath -PathType Leaf) {
      [void]$bins.Add($binPath)
    }
  }

  return , @([System.Collections.Generic.List[string]]::new($bins).ToArray())
}

function Get-GdiReferencedDataFiles {
  param([Parameter(Mandatory = $true)][string]$gdiPath)
  $gdiDir = Split-Path -Parent $gdiPath
  $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $list = [System.Collections.Generic.List[string]]::new()
  $allLines = @(Get-Content -LiteralPath $gdiPath -ErrorAction SilentlyContinue)
  $startIdx = 0
  if ($allLines.Count -gt 0 -and $allLines[0].Trim() -match '^\d+$') {
    $startIdx = 1
  }
  for ($ix = $startIdx; $ix -lt $allLines.Count; $ix++) {
    $line = $allLines[$ix]
    $t = $line.Trim()
    if ($t.Length -eq 0 -or $t.StartsWith('#')) { continue }
    $got = $false
    $parts = @($t -split '\s+')
    if ($parts.Count -ge 5) {
      $name = $parts[4].Trim('"').Trim("'")
      if ($name -match '(?i)\.(bin|raw|iso|img)$') {
        $full = if ([IO.Path]::IsPathRooted($name)) { $name } else { Join-Path $gdiDir $name }
        try { $full = [IO.Path]::GetFullPath($full) } catch { }
        if ((Test-Path -LiteralPath $full -PathType Leaf) -and $seen.Add($full)) {
          [void]$list.Add($full)
          $got = $true
        }
      }
    }
    if (-not $got) {
      foreach ($m in [regex]::Matches($line, '(?i)\b([a-z0-9][a-z0-9._\-]*\.(?:bin|raw|iso|img))\b')) {
        $name = $m.Groups[1].Value
        $full = Join-Path $gdiDir $name
        try { $full = [IO.Path]::GetFullPath($full) } catch { }
        if ((Test-Path -LiteralPath $full -PathType Leaf) -and $seen.Add($full)) {
          [void]$list.Add($full)
        }
      }
    }
  }
  return , @($list.ToArray())
}

function Get-DiscImageFootprintBytes {
  param(
    [Parameter(Mandatory = $true)][string]$LiteralPath,
    [Parameter(Mandatory = $true)][string]$ExtLower
  )
  if (-not (Test-Path -LiteralPath $LiteralPath -PathType Leaf)) {
    return [long]0
  }
  $sum = [long](Get-Item -LiteralPath $LiteralPath).Length
  if ($ExtLower -eq '.cue') {
    foreach ($b in @(Get-CueReferencedBins -cuePath $LiteralPath)) {
      if (Test-Path -LiteralPath $b -PathType Leaf) {
        $sum += [long](Get-Item -LiteralPath $b).Length
      }
    }
  }
  elseif ($ExtLower -eq '.gdi') {
    foreach ($gf in @(Get-GdiReferencedDataFiles -gdiPath $LiteralPath)) {
      if (Test-Path -LiteralPath $gf -PathType Leaf) {
        $sum += [long](Get-Item -LiteralPath $gf).Length
      }
    }
  }
  return $sum
}

# Sum on-disk bytes of all files under a folder that are not .zip/.7z/.rar (avoids counting nested archives as "original" payload).
function Get-TotalNonArchiveFileBytesUnderDirectory {
  param([Parameter(Mandatory = $true)][string]$LiteralPath)
  if (-not (Test-Path -LiteralPath $LiteralPath -PathType Container)) {
    return [long]0
  }
  [long]$sum = 0
  foreach ($f in @(Get-ChildItem -LiteralPath $LiteralPath -Recurse -File -ErrorAction SilentlyContinue)) {
    if (Test-IsArchiveFileName -FileName $f.Name) { continue }
    $sum += [long]$f.Length
  }
  return $sum
}

function Update-ChdmaidExtractDirRootsOrder {
  if ($null -eq $script:chdmaidExtractDirRoots -or $script:chdmaidExtractDirRoots.Count -le 1) { return }
  $script:chdmaidExtractDirRoots.Sort({
      param([string]$a, [string]$b)
      $d = $b.Length - $a.Length
      if ($d -ne 0) { return $d }
      return [string]::Compare($a, $b, [StringComparison]::OrdinalIgnoreCase)
    })
}

function Test-ChdmaidInputUnderKnownExtractDir {
  param([Parameter(Mandatory = $true)][string]$LiteralPath)
  if ($null -eq $script:chdmaidExtractDirRoots -or $script:chdmaidExtractDirRoots.Count -eq 0) {
    return $false
  }
  try {
    $norm = [IO.Path]::GetFullPath($LiteralPath)
  } catch {
    return $false
  }
  $sep = [IO.Path]::DirectorySeparatorChar
  foreach ($root in $script:chdmaidExtractDirRoots) {
    $r = $root.TrimEnd('\').TrimEnd('/')
    if ($norm.Equals($r, [StringComparison]::OrdinalIgnoreCase) -or
        $norm.StartsWith($r + $sep, [StringComparison]::OrdinalIgnoreCase)) {
      return $true
    }
  }
  return $false
}

function Get-ChdmaidLongestExtractRootForPath {
  param([Parameter(Mandatory = $true)][string]$LiteralPath)
  if ($null -eq $script:chdmaidExtractDirRoots -or $script:chdmaidExtractDirRoots.Count -eq 0) { return $null }
  try {
    $norm = [IO.Path]::GetFullPath($LiteralPath)
  } catch {
    return $null
  }
  $sep = [IO.Path]::DirectorySeparatorChar
  [string]$best = $null
  foreach ($root in $script:chdmaidExtractDirRoots) {
    $r = [IO.Path]::GetFullPath($root.TrimEnd('\').TrimEnd('/'))
    if ($norm.Equals($r, [StringComparison]::OrdinalIgnoreCase) -or
        $norm.StartsWith($r + $sep, [StringComparison]::OrdinalIgnoreCase)) {
      if ($null -eq $best -or $r.Length -gt $best.Length) { $best = $r }
    }
  }
  if ($null -eq $best) { return $null }
  return [IO.Path]::GetFullPath($best)
}

# Record on-disk source material for loose paths (not under archive extract dirs): full game folder once, or disc footprint when the image lives directly under -source.
function Register-ChdmaidLooseFolderOriginalPayload {
  param([Parameter(Mandatory = $true)][string]$LiteralPath)
  if (Test-ChdmaidInputUnderKnownExtractDir -LiteralPath $LiteralPath) { return }
  try {
    $srcN = [IO.Path]::GetFullPath($source).TrimEnd([char[]]@('\', '/'))
    $par = [IO.Path]::GetFullPath((Split-Path -Parent $LiteralPath))
  } catch {
    return
  }
  $sep = [IO.Path]::DirectorySeparatorChar
  if ($par.Equals($srcN, [StringComparison]::OrdinalIgnoreCase)) {
    if (-not $script:chdmaidLooseRootDiscPayloadByPath.ContainsKey($LiteralPath)) {
      $extL = [IO.Path]::GetExtension($LiteralPath).ToLowerInvariant()
      $script:chdmaidLooseRootDiscPayloadByPath[$LiteralPath] = Get-DiscImageFootprintBytes -LiteralPath $LiteralPath -ExtLower $extL
    }
    return
  }
  if (-not ($par.StartsWith($srcN + $sep, [StringComparison]::OrdinalIgnoreCase))) { return }
  if ($script:chdmaidLooseFolderPayloadByDir.ContainsKey($par)) { return }
  if (-not (Test-Path -LiteralPath $par -PathType Container)) { return }
  $script:chdmaidLooseFolderPayloadByDir[$par] = Get-TotalNonArchiveFileBytesUnderDirectory -LiteralPath $par
}

$deleteText = $(if ($deleteSources) { 'Yes (-yes)' } else { 'No' })
# Align value column (widen for longest status label)
$labelPadWidth = [Math]::Max([Math]::Max('Source:'.Length, 'Destination:'.Length), [Math]::Max('Delete sources:'.Length, 'Ignore sources:'.Length))

$script:chdmaidIgnoreExistingDest = [bool]$ignore
if (
  (-not $PSBoundParameters.ContainsKey('ignore') -and -not $PSBoundParameters.ContainsKey('yes')) -and
  (@($existingChdsInDest).Count -gt 0)
) {
  Write-Host '.chd files already exist under the destination folder:' -ForegroundColor DarkYellow
  Write-Host '  [I] Ignore  - Do not verify these files, (same as -ignore).' -ForegroundColor White
  Write-Host '  [R] Rescan  - Verify the existing .chd; delete and recreate if invalid.' -ForegroundColor White
  $choice = Read-Host 'Choice [I/R] (Enter = Rescan)'
  $t = if ($null -eq $choice) { '' } else { $choice.Trim() }
  if ($t.Length -gt 0 -and ($t.Substring(0, 1) -ieq 'i')) {
    $script:chdmaidIgnoreExistingDest = $true
  } else {
    $script:chdmaidIgnoreExistingDest = $false
  }
  Write-Host ""
}

$ignoreSourcesText = if ($script:chdmaidIgnoreExistingDest) { 'Yes (-ignore)' } else { 'No (rescan)' }
Write-Host ("{0,-$labelPadWidth}" -f 'Source:') -NoNewline -ForegroundColor DarkCyan
Write-Host (' ' + $source) -ForegroundColor White

Write-Host ("{0,-$labelPadWidth}" -f 'Destination:') -NoNewline -ForegroundColor DarkCyan
Write-Host (' ' + $dest) -ForegroundColor White

Write-Host ("{0,-$labelPadWidth}" -f 'Delete sources:') -NoNewline -ForegroundColor DarkCyan
Write-Host (' ' + $deleteText) -ForegroundColor White

Write-Host ("{0,-$labelPadWidth}" -f 'Ignore sources:') -NoNewline -ForegroundColor DarkCyan
Write-Host (' ' + $ignoreSourcesText) -ForegroundColor White
Write-Host ""

# One recurse over -source: count archives, collect loose .cue/.gdi/.iso (archive contents not listed until extract).
$inputsPre = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
$inputSeenPre = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
[int]$sourceArchiveCount = 0
foreach ($f in Get-ChildItem -LiteralPath $source -Recurse -File -ErrorAction SilentlyContinue) {
  $leafName = $f.Name
  if (Test-IsArchiveFileName -FileName $leafName) {
    $sourceArchiveCount++
    continue
  }
  if (Test-IsDiscInputFileName -FileName $leafName) {
    if ($inputSeenPre.Add($f.FullName)) {
      $inputsPre.Add($f) | Out-Null
    }
  }
}
$hasSourceArchives = ($sourceArchiveCount -gt 0)
$inputsPre = @($inputsPre | Sort-Object FullName)
if ($inputsPre -isnot [System.Array]) {
  $inputsPre = @($inputsPre)
}

if ((@($inputsPre).Count -eq 0) -and (-not $hasSourceArchives)) {
  Write-Warn "No .cue/.gdi/.iso inputs were found under: $source (including extracted archive folders there)."
  exit 0
}

$flags = @()
if ($yes) { $flags += '-yes' }
elseif ($no) { $flags += '-no' }
if ($nolog) { $flags += '-nolog' }
if ($script:chdmaidIgnoreExistingDest) { $flags += '-ignore' }
$quickCmd = '.\chd-maid.ps1 -source ' + (Get-DoubleQuotedPath $source) + ' -dest ' + (Get-DoubleQuotedPath $dest) + ' ' + ($flags -join ' ')
if ($PSBoundParameters.ContainsKey('SevenZipPath') -and -not [string]::IsNullOrWhiteSpace($SevenZipPath)) {
  $quickCmd += ' -SevenZipPath ' + (Get-DoubleQuotedPath $SevenZipPath)
}
$quickCmdLine = $quickCmd.Trim()

# Log files under .\logs\ from the current directory (folder created if missing)
$logsDir = Join-Path (Get-Location).Path 'logs'
$null = New-Item -ItemType Directory -Path $logsDir -Force
$logStamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
$logPath = Join-Path $logsDir ("log-chd-maid-{0}.log" -f $logStamp)
$script:chdMaidErrorLogInitialized = $false

function Initialize-ChdMaidErrorOnlyLog {
  param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$QuickCmdLine,
    [Parameter(Mandatory = $true)][string]$CompatibleDiscoveryLine,
    [Parameter(Mandatory = $true)][int]$ExistingChdCount
  )
  if ($script:chdMaidErrorLogInitialized) { return }
  $script:chdMaidErrorLogInitialized = $true
  @(
    $QuickCmdLine
    ''
    '(-nolog: log created because at least one conversion failed.)'
    ''
    $CompatibleDiscoveryLine
    ("Found $ExistingChdCount existing .chd file(s) under destination.")
    ''
  ) | Set-Content -LiteralPath $Path -Encoding utf8
}

# Pre-extraction headline: loose .cue/.gdi/.iso on disk, plus one nominal slot per archive (contents are not listed until after 7-Zip).
$displayCompatCount = if ($sourceArchiveCount -gt 0) { @($inputsPre).Count + $sourceArchiveCount } else { @($inputsPre).Count }
$lineFoundCompat = "Found $displayCompatCount compatible file(s) in the source folder"
if ($sourceArchiveCount -gt 0) {
  if ($sourceArchiveCount -eq 1) {
    $lineFoundCompat += ', which includes 1 archive file for processing.'
  } else {
    $lineFoundCompat += ", which includes $sourceArchiveCount archive files for processing."
  }
} else {
  $lineFoundCompat += '.'
}

Write-Host "Start Command: " -NoNewline -ForegroundColor Yellow
Write-Host $quickCmdLine -ForegroundColor DarkGray
Write-Host ""

Write-Host "Found " -NoNewline -ForegroundColor White
Write-Host $displayCompatCount -NoNewline -ForegroundColor Green
Write-Host " compatible file(s) in the source folder" -NoNewline -ForegroundColor White
if ($sourceArchiveCount -gt 0) {
  if ($sourceArchiveCount -eq 1) {
    Write-Host ", which includes " -NoNewline -ForegroundColor White
    Write-Host "1" -NoNewline -ForegroundColor Green
    Write-Host " archive file for processing." -ForegroundColor White
  } else {
    Write-Host ", which includes " -NoNewline -ForegroundColor White
    Write-Host $sourceArchiveCount -NoNewline -ForegroundColor Green
    Write-Host " archive files for processing." -ForegroundColor White
  }
} else {
  Write-Host "." -ForegroundColor White
}

Write-Host "Found " -NoNewline -ForegroundColor White
Write-Host @($existingChdsInDest).Count -NoNewline -ForegroundColor Green
Write-Host " existing .chd file(s) under destination.`n" -ForegroundColor White

if ($writeFullLog) {
  @(
    $quickCmdLine
    ''
    $lineFoundCompat
    ("Found $(@($existingChdsInDest).Count) existing .chd file(s) under destination.")
    ''
  ) | Set-Content -LiteralPath $logPath -Encoding utf8
}

function Remove-ChdMaidSourceMediaIfYes {
  param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$ExtLower
  )
  if (-not $deleteSources) { return }
  $toDelete = [System.Collections.Generic.List[string]]::new()
  $toDelete.Add($InputPath) | Out-Null
  if ($ExtLower -eq '.cue') {
    foreach ($b in @(Get-CueReferencedBins -cuePath $InputPath)) {
      $toDelete.Add($b) | Out-Null
    }
  }
  elseif ($ExtLower -eq '.gdi') {
    foreach ($gf in @(Get-GdiReferencedDataFiles -gdiPath $InputPath)) {
      $toDelete.Add($gf) | Out-Null
    }
  }
  foreach ($p in ($toDelete | Sort-Object -Unique)) {
    if (Test-Path -LiteralPath $p -PathType Leaf) {
      Write-Host "Deleting source: " -NoNewline -ForegroundColor White
      Write-Host $p -ForegroundColor DarkGray
      Add-ChdmaidPerFileLogLine -Line ('Deleting source: ' + $p)
      Remove-Item -LiteralPath $p -Force -ErrorAction Stop
    }
  }

  try {
    $sourceNorm = [IO.Path]::GetFullPath($source).TrimEnd([char[]]@('\', '/'))
  } catch {
    return
  }
  $sep = [IO.Path]::DirectorySeparatorChar

  # Remove the whole per-game folder when nothing is left that still needs converting (no .cue/.gdi/.iso and no archives). Applies under loose trees and under archive extract trees (not only empty-dir pruning).
  try {
    $par = [IO.Path]::GetFullPath((Split-Path -Parent $InputPath))
  } catch {
    $par = $null
  }
  if (-not [string]::IsNullOrWhiteSpace($par) -and
    (-not $par.Equals($sourceNorm, [StringComparison]::OrdinalIgnoreCase)) -and
    $par.StartsWith($sourceNorm + $sep, [StringComparison]::OrdinalIgnoreCase) -and
    (Test-Path -LiteralPath $par -PathType Container)) {
    $blocking = @(Get-ChildItem -LiteralPath $par -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { (Test-IsDiscInputFileName -FileName $_.Name) -or (Test-IsArchiveFileName -FileName $_.Name) })
    if ($blocking.Count -eq 0) {
      Write-Host "Removing game folder: " -NoNewline -ForegroundColor DarkCyan
      Write-Host $par -ForegroundColor DarkGray
      Add-ChdmaidPerFileLogLine -Line ('Removing game folder: ' + $par)
      $script:chdmaidSectionSeparatorBeforeNextItem = $true
      Remove-Item -LiteralPath $par -Recurse -Force -ErrorAction Stop
      return
    }
  }

  # Under extract trees (or multiple disc sets in one folder): only prune now-empty parents.
  $folderCandidates = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($fp in @($toDelete | Sort-Object -Unique)) {
    $par = Split-Path -Parent $fp
    while (-not [string]::IsNullOrWhiteSpace($par)) {
      try {
        $norm = [IO.Path]::GetFullPath($par)
      } catch {
        break
      }
      if ($norm.Equals($sourceNorm, [StringComparison]::OrdinalIgnoreCase)) { break }
      if (-not ($norm.StartsWith($sourceNorm + $sep, [StringComparison]::OrdinalIgnoreCase))) { break }
      [void]$folderCandidates.Add($norm)
      $par = Split-Path -Parent $par
    }
  }
  foreach ($d in ($folderCandidates | Sort-Object { $_.Length } -Descending)) {
    if (-not (Test-Path -LiteralPath $d -PathType Container)) { continue }
    $items = @(Get-ChildItem -LiteralPath $d -Force -ErrorAction SilentlyContinue)
    if ($items.Count -ne 0) { continue }
    Write-Host "Removing empty folder: " -NoNewline -ForegroundColor DarkCyan
    Write-Host $d -ForegroundColor DarkGray
    Add-ChdmaidPerFileLogLine -Line ('Removing empty folder: ' + $d)
    Remove-Item -LiteralPath $d -Force -ErrorAction Stop
  }
}

function Invoke-ChdMaidSingleDiscInput {
  param(
    [Parameter(Mandatory = $true)][System.IO.FileInfo]$Item
  )
  [string]$curParent = $null
  try {
    # Blank line after archive / folder-removal at source walk depth, and between disc jobs in different folders (not multi-disc .cue sets in the same folder).
    [bool]$printedLeadingBlank = $false
    if ($script:chdmaidPostExtractWalkDepth -eq 0 -and $script:chdmaidSectionSeparatorBeforeNextItem) {
      Write-Host ''
      $script:chdmaidSectionSeparatorBeforeNextItem = $false
      $printedLeadingBlank = $true
    }
    try {
      $curParent = [IO.Path]::GetFullPath((Split-Path -Parent $Item.FullName))
    } catch {
      $curParent = $null
    }

    $script:chdmaidPerFileLogLines = [System.Collections.Generic.List[string]]::new()
    $script:chdmaidDiscJobIndex++
    if ($script:chdmaidDiscJobIndex -gt 1) {
      $parentChanged = $true
      if ($null -ne $curParent -and $null -ne $script:chdmaidPrevDiscInputParent) {
        $parentChanged = -not $curParent.Equals($script:chdmaidPrevDiscInputParent, [StringComparison]::OrdinalIgnoreCase)
      }
      if ($parentChanged -and -not $printedLeadingBlank) {
        Write-Host ''
      }
    }
    if ($script:chdmaidDiscJobIndex -gt $script:chdmaidDiscJobTotal) {
      $script:chdmaidDiscJobTotal = $script:chdmaidDiscJobIndex
    }
    if ($script:chdmaidDiscJobTotal -gt 0) {
      $pCur = Format-UsInt32 -Value $script:chdmaidDiscJobIndex
      $pTot = Format-UsInt32 -Value $script:chdmaidDiscJobTotal
      $progLine = ('===[ {0}/{1} ]===' -f $pCur, $pTot)
      Write-Host '===[ ' -NoNewline -ForegroundColor DarkCyan
      Write-Host $pCur -NoNewline -ForegroundColor Green
      Write-Host '/' -NoNewline -ForegroundColor DarkCyan
      Write-Host $pTot -NoNewline -ForegroundColor Green
      Write-Host ' ]===' -ForegroundColor DarkCyan
      Add-ChdmaidPerFileLogLine -Line $progLine
    }
    $inputPath = $Item.FullName
    $ext = $Item.Extension.ToLowerInvariant()
    $baseName = [IO.Path]::GetFileNameWithoutExtension($inputPath)
    $outPath = Join-Path $dest ($baseName + ".chd")
    try {
    $recreateAfterRemovingBadChd = $false
    if ($script:chdmaidIgnoreExistingDest -and (Test-Path -LiteralPath $outPath -PathType Leaf)) {
      Write-Host "Ignoring (destination CHD exists): " -NoNewline -ForegroundColor DarkCyan
      Write-Host $outPath -ForegroundColor DarkGray
      Write-Host "  Source: " -NoNewline -ForegroundColor DarkCyan
      Write-Host $inputPath -ForegroundColor DarkGray
      if ($script:chdmaidDiscJobTotal -gt 0) {
        Add-ChdmaidPerFileLogLine -Line ('===[ {0}/{1} ]===' -f (Format-UsInt32 -Value $script:chdmaidDiscJobIndex), (Format-UsInt32 -Value $script:chdmaidDiscJobTotal))
      }
      Add-ChdmaidPerFileLogLine -Line ("IGNORING (destination CHD exists)    {0}" -f $outPath)
      Add-ChdmaidPerFileLogLine -Line $inputPath
      $script:inputOutcome[$inputPath] = 'skipped'
      $script:chdmaidDiscSkippedIgnore++
      Register-ChdmaidLooseFolderOriginalPayload -LiteralPath $inputPath
      Remove-ChdMaidSourceMediaIfYes -InputPath $inputPath -ExtLower $ext
      Complete-ChdmaidPerFileLog -Persist $writeFullLog
      return
    }
    if (Test-Path -LiteralPath $outPath) {
      $fileNameChd = [IO.Path]::GetFileName($outPath)
      $verifyExistingStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
      $verifyExistingResult = Invoke-ProcessWithStatus -ExePath $chdmanExe -Arguments @('verify', '-i', $outPath) -Label 'Verifying' -FileName $fileNameChd -Stopwatch $verifyExistingStopwatch
      Write-Host ""
      Add-ChdmaidPerFileLogLine -Line ''
      if ($verifyExistingResult.ExitCode -eq 0) {
        Write-Host "Skipping (already existed and valid): " -NoNewline -ForegroundColor DarkCyan
        Write-Host $outPath -ForegroundColor DarkGray
        Add-ChdmaidPerFileLogLine -Line $inputPath
        Add-ChdmaidPerFileLogLine -Line ("SKIPPED (already existed and valid)    {0}" -f $outPath)
        $script:inputOutcome[$inputPath] = 'skipped'
        $script:skipped++
        Register-ChdmaidLooseFolderOriginalPayload -LiteralPath $inputPath
        Remove-ChdMaidSourceMediaIfYes -InputPath $inputPath -ExtLower $ext
        Complete-ChdmaidPerFileLog -Persist $writeFullLog
        return
      }
      Write-Host "Existing CHD failed verify; removing defective file: " -NoNewline -ForegroundColor Red
      Write-Host $outPath -ForegroundColor DarkGray
      Add-ChdmaidPerFileLogLine -Line ('Existing CHD failed verify; removing defective file: ' + $outPath)
      if (Test-Path -LiteralPath $outPath) {
        Remove-Item -LiteralPath $outPath -Force -ErrorAction Stop
      }
      Add-ChdmaidPerFileLogLine -Line ("REMOVED (failed verify)    {0}" -f $outPath)
      $recreateAfterRemovingBadChd = $true
    }

    Write-Host "Creating CHD from: " -NoNewline -ForegroundColor White
    Write-Host $inputPath -ForegroundColor DarkGray
    Add-ChdmaidPerFileLogLine -Line ("Creating CHD from: " + $inputPath)
    $fileNameOnly = [IO.Path]::GetFileName($inputPath)
    $createResult = Invoke-ProcessWithStatus -ExePath $chdmanExe -Arguments @('createcd','-i',$inputPath,'-o',$outPath) -Label 'Creating' -FileName $fileNameOnly -Stopwatch ([System.Diagnostics.Stopwatch]::StartNew())
    Write-Host ""
    Add-ChdmaidPerFileLogLine -Line ''

    if ($createResult.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $outPath)) {
      throw "chdman createcd failed."
    }

    $verifyResult = Invoke-ProcessWithStatus -ExePath $chdmanExe -Arguments @('verify','-i',$outPath) -Label 'Verifying' -FileName $fileNameOnly -Stopwatch ([System.Diagnostics.Stopwatch]::StartNew())
    Write-Host ""
    Add-ChdmaidPerFileLogLine -Line ''
    if ($verifyResult.ExitCode -ne 0) {
      if (Test-Path -LiteralPath $outPath -PathType Leaf) {
        Write-Host "CHD failed verify; removing defective file: " -NoNewline -ForegroundColor Red
        Write-Host $outPath -ForegroundColor DarkGray
        Add-ChdmaidPerFileLogLine -Line ('CHD failed verify; removing defective file: ' + $outPath)
        Remove-Item -LiteralPath $outPath -Force -ErrorAction Stop
      }
      throw "chdman verify failed after CHD was successfully created."
    }

    Write-Host "Verified: " -NoNewline -ForegroundColor DarkCyan
    Write-Host $outPath -ForegroundColor DarkGray
    Add-ChdmaidPerFileLogLine -Line $inputPath
    Add-ChdmaidPerFileLogLine -Line ("VERIFIED    {0}" -f $outPath)
    if ($recreateAfterRemovingBadChd) {
      $script:chdmaidChdsRecreated++
    }
    $script:inputOutcome[$inputPath] = 'success'
    $script:bytesChdThisRun += [long](Get-Item -LiteralPath $outPath).Length
    $script:success++
    Register-ChdmaidLooseFolderOriginalPayload -LiteralPath $inputPath

    Remove-ChdMaidSourceMediaIfYes -InputPath $inputPath -ExtLower $ext
    Complete-ChdmaidPerFileLog -Persist $writeFullLog
  } catch {
    $script:failed++
    $script:inputOutcome[$inputPath] = 'failed'
    if (Test-Path -LiteralPath $outPath -PathType Leaf) {
      Write-Host "Removing failed or partial CHD output: " -NoNewline -ForegroundColor Red
      Write-Host $outPath -ForegroundColor DarkGray
      Add-ChdmaidPerFileLogLine -Line ('Removing failed or partial CHD output: ' + $outPath)
      try {
        Remove-Item -LiteralPath $outPath -Force -ErrorAction Stop
      } catch {
        Write-Warn ("Could not remove failed CHD output: {0} | {1}" -f $outPath, $_.Exception.Message)
      }
    }
    Write-Host "Failed: " -NoNewline -ForegroundColor Red
    Write-Host $inputPath -ForegroundColor DarkGray
    Add-ChdmaidPerFileLogLine -Line ('Failed: ' + $inputPath)
    if ($_.Exception -and $_.Exception.Message) {
      Write-Host ("  {0}" -f $_.Exception.Message) -ForegroundColor DarkRed
      Add-ChdmaidPerFileLogLine -Line ('  ' + $_.Exception.Message)
    }
    $failDetail = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { 'Unknown error' }
    if (-not $writeFullLog) {
      $showErrorLogPath = -not $script:chdMaidErrorLogInitialized
      Initialize-ChdMaidErrorOnlyLog -Path $logPath -QuickCmdLine $quickCmdLine -CompatibleDiscoveryLine $lineFoundCompat -ExistingChdCount @($existingChdsInDest).Count
      if ($showErrorLogPath) {
        Write-Host "Error log: " -NoNewline -ForegroundColor Yellow
        Write-Host $logPath -ForegroundColor DarkGray
      }
    }
    Add-ChdmaidPerFileLogLine -Line ("FAILED    {0} | {1}" -f $outPath, $failDetail)
    Complete-ChdmaidPerFileLog -Persist $true
  }
  } finally {
    $script:chdmaidPrevDiscInputParent = $curParent
  }
}

function Expand-OneArchiveAndVisit {
  param(
    [Parameter(Mandatory = $true)][System.IO.FileInfo]$ArchiveFileInfo,
    [Parameter(Mandatory = $true)][string]$SourceRootFull,
    [Parameter(Mandatory = $true)][string]$SevenZipExe,
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$ArchiveRecords,
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.HashSet[string]]$DoneArchives
  )
  if ([string]::IsNullOrWhiteSpace($SevenZipExe)) {
    throw '7-Zip executable path is missing; archives cannot be extracted.'
  }
  if (-not $DoneArchives.Add($ArchiveFileInfo.FullName)) { return }
  if ($script:chdmaidIgnoreExistingDest) {
    $stemIgn = Get-ArchiveStemForFolderName -FileName $ArchiveFileInfo.Name
    $destChdFromArchiveName = Join-Path $dest ($stemIgn + '.chd')
    if (Test-Path -LiteralPath $destChdFromArchiveName -PathType Leaf) {
      if ($script:chdmaidSectionSeparatorBeforeNextItem) {
        Write-Host ''
        $script:chdmaidSectionSeparatorBeforeNextItem = $false
      }
      Write-Host "Ignoring archive (-ignore; destination CHD exists): " -NoNewline -ForegroundColor DarkCyan
      Write-Host $ArchiveFileInfo.FullName -ForegroundColor DarkGray
      Write-Host "  -> " -NoNewline -ForegroundColor DarkCyan
      Write-Host $destChdFromArchiveName -ForegroundColor DarkGray
      $script:chdmaidArchivesSkippedIgnore++
      $script:chdmaidArchiveExtractPayloadByDir[$ArchiveFileInfo.FullName] = [long](Get-Item -LiteralPath $ArchiveFileInfo.FullName -ErrorAction Stop).Length
      if ($script:chdmaidDiscJobTotal -gt $script:chdmaidDiscJobIndex) {
        $script:chdmaidDiscJobTotal--
      }
      if ($writeFullLog) {
        Add-Content -LiteralPath $logPath -Encoding utf8 -Value @(
          "IGNORED (archive; -ignore; CHD in destination)    $($ArchiveFileInfo.FullName)",
          $destChdFromArchiveName,
          ''
        )
      }
      $script:chdmaidSectionSeparatorBeforeNextItem = $true
      return
    }
  }
  if ($script:chdmaidSectionSeparatorBeforeNextItem) {
    Write-Host ''
    $script:chdmaidSectionSeparatorBeforeNextItem = $false
  }
  $stem = Get-ArchiveStemForFolderName -FileName $ArchiveFileInfo.Name
  $extractDir = Get-UniqueExtractDirectoryPath -ParentRoot $SourceRootFull -FolderStem $stem
  Write-Host "Extracting archive: " -NoNewline -ForegroundColor White
  Write-Host $ArchiveFileInfo.FullName -ForegroundColor DarkGray
  Write-Host "-> " -NoNewline -ForegroundColor DarkCyan
  Write-Host $extractDir -ForegroundColor DarkGray

  $script:chdmaidPostExtractWalkDepth++
  $archivePathForCleanup = $ArchiveFileInfo.FullName
  $archiveExtractFinished = $false
  try {
    Expand-ArchiveWith7Zip -LiteralArchivePath $ArchiveFileInfo.FullName -ExtractDir $extractDir -SevenZipExe $SevenZipExe
    $archiveExtractFinished = $true
    [void]$ArchiveRecords.Add([pscustomobject]@{ ExtractDir = $extractDir; ArchivePath = $ArchiveFileInfo.FullName })
    $extractNorm = [IO.Path]::GetFullPath($extractDir)
    [void]$script:chdmaidExtractDirRoots.Add($extractNorm)
    Update-ChdmaidExtractDirRootsOrder
    # Original-size stats use the archive file on disk, not the extracted folder tree.
    $script:chdmaidArchiveExtractPayloadByDir[$extractNorm] = [long](Get-Item -LiteralPath $ArchiveFileInfo.FullName -ErrorAction Stop).Length

    Invoke-ChdMaidSourceDirectory -LiteralDir $extractDir -SourceRootFull $SourceRootFull -SevenZipExe $SevenZipExe -ArchiveRecords $ArchiveRecords -DoneArchives $DoneArchives
  } finally {
    # Remove unpack tree (and archive file if -yes) before returning to the parent scan so the next game never sees leftover extract data.
    if ($archiveExtractFinished) {
      Remove-ArchiveExtractionIfReady -ExtractDir $extractDir -ArchivePath $archivePathForCleanup -InputOutcome $script:inputOutcome -AllInputPaths @([string[]]@($script:inputOutcome.Keys))
    }
    if ($script:chdmaidPostExtractWalkDepth -gt 0) {
      $script:chdmaidPostExtractWalkDepth--
    }
    $script:chdmaidSectionSeparatorBeforeNextItem = $true
  }
}

function Invoke-ChdMaidSourceDirectory {
  param(
    [Parameter(Mandatory = $true)][string]$LiteralDir,
    [Parameter(Mandatory = $true)][string]$SourceRootFull,
    [AllowNull()]
    [AllowEmptyString()]
    [string]$SevenZipExe,
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$ArchiveRecords,
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.HashSet[string]]$DoneArchives
  )
  if (-not (Test-Path -LiteralPath $LiteralDir -PathType Container)) { return }
  $children = @(Get-ChildItem -LiteralPath $LiteralDir -Force -ErrorAction SilentlyContinue | Sort-Object FullName)
  foreach ($ch in $children) {
    if ($ch.PSIsContainer) {
      Invoke-ChdMaidSourceDirectory -LiteralDir $ch.FullName -SourceRootFull $SourceRootFull -SevenZipExe $SevenZipExe -ArchiveRecords $ArchiveRecords -DoneArchives $DoneArchives
    } elseif (Test-IsArchiveFileName -FileName $ch.Name) {
      Expand-OneArchiveAndVisit -ArchiveFileInfo $ch -SourceRootFull $SourceRootFull -SevenZipExe $SevenZipExe -ArchiveRecords $ArchiveRecords -DoneArchives $DoneArchives
    } elseif (Test-IsDiscInputFileName -FileName $ch.Name) {
      Invoke-ChdMaidSingleDiscInput -Item $ch
    }
  }
}

$overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$sevenZipExeResolved = $null
if ($hasSourceArchives) {
  try {
    $sevenZipExeResolved = Resolve-SevenZipExecutable -BoundValue $SevenZipPath -WasBound $PSBoundParameters.ContainsKey('SevenZipPath') -ExecutedDir $executedDir
  } catch {
    $failMsg = $_.Exception.Message
    if ([string]::IsNullOrWhiteSpace($failMsg)) {
      $failMsg = $_.ToString()
    }
    if ([string]::IsNullOrWhiteSpace($failMsg)) {
      $failMsg = 'An error occurred while resolving the 7-Zip executable path.'
    }
    Write-Fail $failMsg
  }
}

$archiveExtractRecords = [System.Collections.Generic.List[object]]::new()
$doneArchives = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$sourceRootFull = [IO.Path]::GetFullPath($source)

$success = 0
$skipped = 0
$failed = 0
$script:chdmaidDiscSkippedIgnore = 0
$script:chdmaidArchivesSkippedIgnore = 0
$script:chdmaidChdsRecreated = 0
$script:inputOutcome = @{}
$script:chdmaidExtractDirRoots = [System.Collections.Generic.List[string]]::new()
$script:chdmaidArchiveExtractPayloadByDir = @{}
$script:chdmaidLooseFolderPayloadByDir = @{}
$script:chdmaidLooseRootDiscPayloadByPath = @{}
[long]$bytesChdThisRun = 0
$script:chdmaidSectionSeparatorBeforeNextItem = $false
$script:chdmaidPostExtractWalkDepth = 0
$script:chdmaidPrevDiscInputParent = $null

# Progress total matches "Found N compatible file(s)" (loose disc images + one slot per archive). If an archive contains more disc jobs than that nominal count, the total grows while running.
$script:chdmaidDiscJobTotal = $displayCompatCount
$script:chdmaidDiscJobIndex = 0

Invoke-ChdMaidSourceDirectory -LiteralDir $sourceRootFull -SourceRootFull $sourceRootFull -SevenZipExe $sevenZipExeResolved -ArchiveRecords $archiveExtractRecords -DoneArchives $doneArchives

if (
  ($success + $skipped + $failed + $script:chdmaidDiscSkippedIgnore + $script:chdmaidArchivesSkippedIgnore) -eq 0
) {
  if ((@($inputsPre).Count -gt 0) -or $hasSourceArchives) {
    Write-Warn "No .cue/.gdi/.iso jobs were completed under: $source."
  }
  exit 0
}

Remove-CompletedArchiveExtractions -ExtractRecords @($archiveExtractRecords) -InputOutcome $script:inputOutcome -AllInputPaths @([string[]]@($script:inputOutcome.Keys))

$overallStopwatch.Stop()

[long]$archivePayloadSum = [long]0
foreach ($kv in $script:chdmaidArchiveExtractPayloadByDir.GetEnumerator()) {
  $archivePayloadSum += [long]$kv.Value
}
[long]$looseFolderPayloadSum = [long]0
foreach ($kv in $script:chdmaidLooseFolderPayloadByDir.GetEnumerator()) {
  $looseFolderPayloadSum += [long]$kv.Value
}
[long]$looseRootPayloadSum = [long]0
foreach ($kv in $script:chdmaidLooseRootDiscPayloadByPath.GetEnumerator()) {
  $looseRootPayloadSum += [long]$kv.Value
}
# Original size: per-archive file on-disk size + each loose game folder once (non-archive files) + loose disc images directly under -source.
[long]$bytesOriginalHandled = $looseFolderPayloadSum + $looseRootPayloadSum + $archivePayloadSum

# Same basis as "Original size" vs "New CHD size" in the completion summary (positive = CHDs smaller than attributed sources).
[long]$bytesSavedNet = $bytesOriginalHandled - $bytesChdThisRun
$archivesExtractedCount = @($archiveExtractRecords).Count

$ts = $overallStopwatch.Elapsed
if ($ts.Days -gt 0) {
  $dayWord = if ($ts.Days -eq 1) { 'day' } else { 'days' }
  $timeText = ('{0} {1}, {2:D2}:{3:D2}:{4:D2}' -f $ts.Days, $dayWord, $ts.Hours, $ts.Minutes, $ts.Seconds)
} else {
  $timeText = ('{0:D2}:{1:D2}:{2:D2}' -f $ts.Hours, $ts.Minutes, $ts.Seconds)
}
$summaryLabelWidth = 38
$summaryLines = [System.Collections.Generic.List[string]]::new()
[void]$summaryLines.Add('')
[void]$summaryLines.Add('===[ Completion Summary ]===')
[void]$summaryLines.Add(("     " + ('Elapsed time:').PadRight($summaryLabelWidth) + $timeText))
if ($success -ne 0) {
  [void]$summaryLines.Add(("     " + ('CHDs created:').PadRight($summaryLabelWidth) + (Format-UsInt32 -Value $success)))
}
if ($script:chdmaidChdsRecreated -ne 0) {
  [void]$summaryLines.Add(("     " + ('CHDs recreated:').PadRight($summaryLabelWidth) + (Format-UsInt32 -Value $script:chdmaidChdsRecreated)))
}
if ($skipped -ne 0) {
  [void]$summaryLines.Add(("     " + ('CHDs skipped (verified existing):').PadRight($summaryLabelWidth) + (Format-UsInt32 -Value $skipped)))
}
if ($script:chdmaidDiscSkippedIgnore -ne 0) {
  [void]$summaryLines.Add(("     " + ('CHDs skipped (-ignore):').PadRight($summaryLabelWidth) + (Format-UsInt32 -Value $script:chdmaidDiscSkippedIgnore)))
}
if ($script:chdmaidArchivesSkippedIgnore -ne 0) {
  [void]$summaryLines.Add(("     " + ('Archives skipped (-ignore):').PadRight($summaryLabelWidth) + (Format-UsInt32 -Value $script:chdmaidArchivesSkippedIgnore)))
}
if ($failed -ne 0) {
  [void]$summaryLines.Add(("     " + ('CHDs failed:').PadRight($summaryLabelWidth) + (Format-UsInt32 -Value $failed)))
}
if ($archivesExtractedCount -ne 0) {
  [void]$summaryLines.Add(("     " + ('Archives extracted:').PadRight($summaryLabelWidth) + (Format-UsInt32 -Value $archivesExtractedCount)))
}
if ($bytesOriginalHandled -ne 0) {
  [void]$summaryLines.Add(("     " + ('Original size (processed disc sets):').PadRight($summaryLabelWidth) + (Format-DataSizeGbMb -Bytes $bytesOriginalHandled)))
}
if ($bytesChdThisRun -ne 0) {
  [void]$summaryLines.Add(("     " + ('New CHD size (this run only):').PadRight($summaryLabelWidth) + (Format-DataSizeGbMb -Bytes $bytesChdThisRun)))
}
if (($bytesOriginalHandled -gt 0) -and ($bytesChdThisRun -gt 0)) {
  $pctSaved = 100.0 * [double]$bytesSavedNet / [double]$bytesOriginalHandled
  $cPct = [cultureinfo]::GetCultureInfo('en-US')
  $pctPart = [string]::Format($cPct, ' - {0:N1}% difference', $pctSaved)
  [void]$summaryLines.Add(("     " + ('Space saved (new CHDs vs source):').PadRight($summaryLabelWidth) + (Format-DataSizeGbMb -Bytes $bytesSavedNet) + $pctPart))
}
$logSummaryLine = $null
if ($writeFullLog) {
  $logSummaryLine = ("     " + ('Log file:').PadRight($summaryLabelWidth) + $logPath)
} elseif (Test-Path -LiteralPath $logPath) {
  $logSummaryLine = ("     " + ('Log file:').PadRight($summaryLabelWidth) + $logPath)
}
if ($null -ne $logSummaryLine) {
  [void]$summaryLines.Add($logSummaryLine)
}

Write-Host ""
Write-Host ""
foreach ($line in $summaryLines | Select-Object -Skip 1) {
  if ($line -match '^(?<prefix>\s+CHDs failed:\s*)(?<num>[\d,]+)\s*$') {
    Write-Host $matches.prefix -NoNewline -ForegroundColor DarkCyan
    Write-Host $matches.num -ForegroundColor DarkRed
  } else {
    Write-Summary $line
  }
}
Write-Host ""

if ($writeFullLog -or $failed -gt 0) {
  Add-Content -LiteralPath $logPath -Value @($summaryLines) -Encoding utf8
}

if ($failed -gt 0) { exit 1 } else { exit 0 }

