[CmdletBinding()]
param(
  [string]$source,
  [string]$dest,
  [switch]$yes,
  [switch]$no,
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
  param([Parameter(Mandatory = $true)][string]$Message)
  Write-Host $Message -ForegroundColor Red
  exit 1
}

function Clear-HostSafe {
  try { Clear-Host } catch { }
}

function Show-Header {
  Write-Host "=== [ CHD Maid ] ===" -ForegroundColor Green
  Write-Host "=== [ Version 2026.3.22 ] ===" -ForegroundColor Yellow
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
  Write-Host ""

  Write-Host "Creates CHD images from `.cue`, `.gdi`, and `.iso` files using `chdman`." -ForegroundColor White
  Write-Info "Requires PowerShell 7.6.0+ (pwsh)."
  Write-Info "By default, source files are NOT deleted."
  Write-Host ""

  Write-Host "Flags:" -ForegroundColor DarkYellow
  Write-MessageWithFlags "  -help      Show this help document."
  Write-MessageWithFlags "  -yes       Delete sources after successful create + `chdman verify`."
  Write-MessageWithFlags "  -no        Do not delete sources (default)."
  Write-Host ""

  Write-Host "Parameters:" -ForegroundColor DarkYellow
  Write-MessageWithFlags "  -source <path>  Source root to scan recursively for `.cue/.gdi/.iso`."
  Write-MessageWithFlags "  -dest   <path>  Destination folder for generated `.chd` files. (default: current folder)"
  Write-Host ""

  Write-Host "Examples:" -ForegroundColor DarkYellow
  Write-MessageWithFlags "  chd-maid.ps1 -source `"D:\source`" -dest `"D:\dest`" -no"
  Write-MessageWithFlags "  chd-maid.ps1 -source `"D:\source`" -dest `"D:\dest`" -yes"
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
  throw "Use only one of -yes or -no."
}

$deleteSources = $false
if ($yes) { $deleteSources = $true }

# When -source / -dest are omitted, prompt until a valid folder exists (same behavior for both).
function Resolve-FolderParameter {
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
    return $BoundValue
  }
  $folder = $null
  while ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
    $folder = Read-Host "$PromptLabel enter to use default (currently: $defaultPath)"
    if ([string]::IsNullOrWhiteSpace($folder)) {
      $folder = $defaultPath
    }
    if (-not (Test-Path -LiteralPath $folder)) {
      Write-Warning "Path not found: $folder"
      $folder = $null
    }
  }
  return $folder
}

$source = Resolve-FolderParameter -PromptLabel 'Source folder' -BoundValue $source -WasBound $PSBoundParameters.ContainsKey('source')
$dest = Resolve-FolderParameter -PromptLabel 'Destination folder' -BoundValue $dest -WasBound $PSBoundParameters.ContainsKey('dest')

Write-Host ""

# Locate chdman.exe
$executedDir = (Get-Location).Path
$chdmanExe = Join-Path $executedDir "chdman.exe"
if (-not (Test-Path -LiteralPath $chdmanExe -PathType Leaf)) {
  Write-Fail "chdman.exe does not exist in the executed directory: $executedDir"
}

New-Item -ItemType Directory -Path $dest -Force | Out-Null

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

function Get-CueReferencedBins {
  param(
    [Parameter(Mandatory = $true)]
    [string]$cuePath
  )

  $cueDir = Split-Path -Parent $cuePath
  $bins = [System.Collections.Generic.List[string]]::new()

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
      if (-not ($bins.Contains($binPath, [StringComparer]::OrdinalIgnoreCase))) {
        $bins.Add($binPath) | Out-Null
      }
    }
  }

  return ,$bins.ToArray()
}

# Scan recursively for inputs (same set as the original .bat)
$inputs = @()
$inputs += Get-ChildItem -Path $source -Recurse -File -Filter "*.cue"
$inputs += Get-ChildItem -Path $source -Recurse -File -Filter "*.gdi"
$inputs += Get-ChildItem -Path $source -Recurse -File -Filter "*.iso"
$inputs = @($inputs | Sort-Object FullName)

if (-not $inputs -or $inputs.Count -eq 0) {
  Write-Warn "No .cue/.gdi/.iso inputs were found under: $source"
  exit 0
}

$existingChdsInDest = @()
if (Test-Path -LiteralPath $dest) {
  $existingChdsInDest = @(Get-ChildItem -LiteralPath $dest -Recurse -File -Filter '*.chd' -ErrorAction SilentlyContinue | Sort-Object FullName)
}

$overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$deleteText = $(if ($deleteSources) { 'Yes (-yes)' } else { 'No' })
# Align value column (longest label is "Delete sources:")
$labelPadWidth = 'Delete sources:'.Length

Write-Host ("{0,-$labelPadWidth}" -f 'Source:') -NoNewline -ForegroundColor DarkCyan
Write-Host (' ' + $source) -ForegroundColor White

Write-Host ("{0,-$labelPadWidth}" -f 'Destination:') -NoNewline -ForegroundColor DarkCyan
Write-Host (' ' + $dest) -ForegroundColor White

Write-Host ("{0,-$labelPadWidth}" -f 'Delete sources:') -NoNewline -ForegroundColor DarkCyan
Write-Host (' ' + $deleteText) -ForegroundColor White
Write-Host ""

$flags = @()
if ($yes) { $flags += '-yes' }
elseif ($no) { $flags += '-no' }
$quickCmd = '.\chd-maid.ps1 -source ' + (Get-DoubleQuotedPath $source) + ' -dest ' + (Get-DoubleQuotedPath $dest) + ' ' + ($flags -join ' ')

# Log file in the directory the script was executed from (current location)
$logStamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
$logPath = Join-Path (Get-Location).Path ("chd-maid-log-{0}.log" -f $logStamp)

Write-Host "Found " -NoNewline -ForegroundColor White
Write-Host $inputs.Count -NoNewline -ForegroundColor Magenta
Write-Host " compatible file(s) in the source folder." -ForegroundColor White

Write-Host "Found " -NoNewline -ForegroundColor White
Write-Host $existingChdsInDest.Count -NoNewline -ForegroundColor Magenta
Write-Host " existing .chd file(s) under destination.`n" -ForegroundColor White

Write-Host "Start Command: " -NoNewline -ForegroundColor Yellow
Write-Host $quickCmd.Trim() -ForegroundColor DarkGray

Write-Host "Log file: " -NoNewline -ForegroundColor Yellow
Write-Host $logPath -ForegroundColor DarkGray
Write-Host "`n"

$quickCmdLine = $quickCmd.Trim()
@(
  $quickCmdLine
  ''
  ("Found $($inputs.Count) compatible file(s).")
  ("Found $($existingChdsInDest.Count) existing .chd file(s) under destination.")
  ''
) | Set-Content -LiteralPath $logPath -Encoding utf8

$success = 0
$skipped = 0
$failed = 0

foreach ($item in $inputs) {
  $inputPath = $item.FullName
  $ext = $item.Extension.ToLowerInvariant()
  $baseName = [IO.Path]::GetFileNameWithoutExtension($inputPath)
  $outPath = Join-Path $dest ($baseName + ".chd")
  try {
    if (Test-Path -LiteralPath $outPath) {
      $fileNameChd = [IO.Path]::GetFileName($outPath)
      $verifyExistingStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
      $verifyExistingResult = Invoke-ProcessWithStatus -ExePath $chdmanExe -Arguments @('verify', '-i', $outPath) -Label 'Verifying' -FileName $fileNameChd -Stopwatch $verifyExistingStopwatch
      Write-Host ""
      if ($verifyExistingResult.ExitCode -eq 0) {
        Write-Host "Skipping (already existed and valid): " -NoNewline -ForegroundColor DarkCyan
        Write-Host $outPath -ForegroundColor DarkGray
        Add-Content -LiteralPath $logPath -Value $inputPath -Encoding utf8
        Add-Content -LiteralPath $logPath -Value ("SKIPPED (already existed and valid)    {0}" -f $outPath) -Encoding utf8
        Add-Content -LiteralPath $logPath -Value '' -Encoding utf8
        $skipped++
        Write-Host ""
        continue
      }
      Write-Host "Existing CHD failed verify; removing defective file: " -NoNewline -ForegroundColor Red
      Write-Host $outPath -ForegroundColor DarkGray
      if (Test-Path -LiteralPath $outPath) {
        Remove-Item -LiteralPath $outPath -Force -ErrorAction Stop
      }
      Add-Content -LiteralPath $logPath -Value ("REMOVED (failed verify)    {0}" -f $outPath) -Encoding utf8
    }

    Write-Host "Creating CHD from: " -NoNewline -ForegroundColor White
    Write-Host $inputPath -ForegroundColor DarkGray
    $fileNameOnly = [IO.Path]::GetFileName($inputPath)
    $createStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $createResult = Invoke-ProcessWithStatus -ExePath $chdmanExe -Arguments @('createcd','-i',$inputPath,'-o',$outPath) -Label 'Creating' -FileName $fileNameOnly -Stopwatch $createStopwatch
    Write-Host ""

    if ($createResult.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $outPath)) {
      throw "chdman createcd failed."
    }

    $verifyStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $verifyResult = Invoke-ProcessWithStatus -ExePath $chdmanExe -Arguments @('verify','-i',$outPath) -Label 'Verifying' -FileName $fileNameOnly -Stopwatch $verifyStopwatch
    Write-Host ""
    if ($verifyResult.ExitCode -ne 0) {
      throw "chdman verify failed after CHD was successfully created."
    }

    Write-Host "Verified: " -NoNewline -ForegroundColor DarkCyan
    Write-Host $outPath -ForegroundColor DarkGray
    Add-Content -LiteralPath $logPath -Value $inputPath -Encoding utf8
    Add-Content -LiteralPath $logPath -Value ("VERIFIED    {0}" -f $outPath) -Encoding utf8
    Add-Content -LiteralPath $logPath -Value '' -Encoding utf8
    $success++

    if ($deleteSources) {
      $toDelete = [System.Collections.Generic.List[string]]::new()
      $toDelete.Add($inputPath) | Out-Null

      if ($ext -eq ".cue") {
        $bins = Get-CueReferencedBins -cuePath $inputPath
        foreach ($b in $bins) {
          $toDelete.Add($b) | Out-Null
        }
      }

      foreach ($p in ($toDelete | Sort-Object -Unique)) {
        if (Test-Path -LiteralPath $p -PathType Leaf) {
          Write-Host "Deleting source: " -NoNewline -ForegroundColor White
          Write-Host $p -ForegroundColor DarkGray
          Remove-Item -LiteralPath $p -Force -ErrorAction Stop
        }
      }
    }

    Write-Host ""
  } catch {
    $failed++
    Write-Host "Failed: " -NoNewline -ForegroundColor Red
    Write-Host $inputPath -ForegroundColor DarkGray
    if ($_.Exception -and $_.Exception.Message) {
      Write-Host ("  {0}" -f $_.Exception.Message) -ForegroundColor DarkRed
    }
    $failDetail = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { 'Unknown error' }
    Add-Content -LiteralPath $logPath -Value $inputPath -Encoding utf8
    Add-Content -LiteralPath $logPath -Value ("FAILED    {0} | {1}" -f $outPath, $failDetail) -Encoding utf8
    Add-Content -LiteralPath $logPath -Value '' -Encoding utf8
  }
}

$overallStopwatch.Stop()

$timeText = $overallStopwatch.Elapsed.ToString('hh\:mm\:ss')
$summaryLabelWidth = 24
$summaryLines = @(
  ''
  '===[ Completion Summary ]==='
  ("     " + ('Elapsed time:').PadRight($summaryLabelWidth) + $timeText)
  ("     " + ('CHDs created:').PadRight($summaryLabelWidth) + $success)
  ("     " + ('CHDs skipped (valid):').PadRight($summaryLabelWidth) + $skipped)
  ("     " + ('CHDs failed:').PadRight($summaryLabelWidth) + $failed)
)

Write-Host ""
foreach ($line in $summaryLines[1..($summaryLines.Length - 1)]) {
  Write-Summary $line
}
Write-Host ""

Add-Content -LiteralPath $logPath -Value $summaryLines -Encoding utf8

if ($failed -gt 0) { exit 1 } else { exit 0 }

