[CmdletBinding()]
param(
  [string]$source,
  [string]$dest,
  [switch]$yes,
  [switch]$no,
  [switch]$help
)

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
  Write-Host "=== [ CHD Maid ] ===" -ForegroundColor DarkGreen
  Write-Host "=== [ Make CHD ] ===" -ForegroundColor DarkGreen
  Write-Host "=== [ Version 2026.3 ] ===" -ForegroundColor DarkYellow
  Write-Host ""
}

# Require PowerShell 7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
  Write-Fail "PowerShell $($PSVersionTable.PSVersion) detected. PowerShell 7 (pwsh) or newer is required."
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

function Write-StatusLine {
  param(
    [Parameter(Mandatory = $true)][string]$Label,
    [Parameter(Mandatory = $true)][string]$FileName,
    [Parameter(Mandatory = $true)][TimeSpan]$Elapsed
  )
  $elapsedText = $Elapsed.ToString('hh\:mm\:ss')
  $progressPrefix = "[$elapsedText] ${Label} - "

  $consoleWidth = 120
  try { $consoleWidth = $Host.UI.RawUI.WindowSize.Width } catch { }
  $maxFileNameLength = [Math]::Max(8, $consoleWidth - $progressPrefix.Length)
  $fileNameDisplay = Format-TextPreview -Text $FileName -MaxLength $maxFileNameLength

  $line = $progressPrefix + $fileNameDisplay
  $pad = ' ' * [Math]::Max(0, $script:lastStatusLineLength - $line.Length)

  Write-Host "`r$progressPrefix" -NoNewline -ForegroundColor White
  Write-Host $fileNameDisplay -NoNewline -ForegroundColor DarkCyan
  Write-Host $pad -NoNewline -ForegroundColor DarkCyan

  $script:lastStatusLineLength = ($line.Length)
}

function Show-Help {
  Clear-HostSafe
  $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
  $scriptName = [IO.Path]::GetFileName($scriptPath)

  Show-Header

  Write-Host "Creates CHD images from `.cue`, `.gdi`, and `.iso` files using `chdman`." -ForegroundColor White
  Write-Info "Requires PowerShell 7+ (pwsh)."
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
  Write-MessageWithFlags "  $scriptName -source `"D:\source`" -dest `"D:\dest`" -no"
  Write-MessageWithFlags "  $scriptName -source `"D:\source`" -dest `"D:\dest`" -yes"
  
  Write-Host ""
}

$helpTokens = @(
  foreach ($a in $args) {
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

function Get-ExistingFolder([string]$promptText, [string]$defaultPath) {
  $folder = $null
  while ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
    $folder = Read-Host "$promptText (default: $defaultPath)"
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

# Only prompt when -source / -dest were not passed on the command line.
if (-not $PSBoundParameters.ContainsKey('source')) {
  $source = Get-ExistingFolder "Source folder for recursive search -" (Get-Location).Path
} elseif ([string]::IsNullOrWhiteSpace($source)) {
  $source = (Get-Location).Path
}

if (-not $PSBoundParameters.ContainsKey('dest')) {
  $defaultDest = (Get-Location).Path
  $dest = Read-Host "Destination folder for .chd files - (default: $defaultDest)"
  if ([string]::IsNullOrWhiteSpace($dest)) {
    $dest = $defaultDest
  }
} elseif ([string]::IsNullOrWhiteSpace($dest)) {
  $dest = (Get-Location).Path
}

# Locate chdman.exe
$executedDir = (Get-Location).Path
$chdmanExe = Join-Path $executedDir "chdman.exe"
if (-not (Test-Path -LiteralPath $chdmanExe -PathType Leaf)) {
  Write-Fail "chdman.exe does not exist in the executed directory: $executedDir"
}

New-Item -ItemType Directory -Path $dest -Force | Out-Null

function Test-ChdmanVerify {
  param(
    [Parameter(Mandatory = $true)]
    [string]$chdPath
  )

  if (-not (Test-Path -LiteralPath $chdPath)) { return $false }

  # chdman verify -i <filename>
  & $chdmanExe verify -i $chdPath 2>&1 | Out-Null
  return ($LASTEXITCODE -eq 0)
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
  [void]$p.Start()

  while (-not $p.HasExited) {
    Write-StatusLine -Label $Label -FileName $FileName -Elapsed $Stopwatch.Elapsed
    Start-Sleep -Milliseconds 250
  }

  # Flush a final update at completion
  Write-StatusLine -Label $Label -FileName $FileName -Elapsed $Stopwatch.Elapsed

  $stdout = $p.StandardOutput.ReadToEnd()
  $stderr = $p.StandardError.ReadToEnd()
  $p.WaitForExit()

  return [pscustomobject]@{
    ExitCode = $p.ExitCode
    StdOut   = $stdout
    StdErr   = $stderr
  }
}

function Get-CueReferencedBins {
  param(
    [Parameter(Mandatory = $true)]
    [string]$cuePath
  )

  $cueDir = Split-Path -Parent $cuePath
  $bins = New-Object System.Collections.Generic.List[string]

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
$overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$labelWidth = 15
$deleteText = $(if ($deleteSources) { 'Yes (-yes)' } else { 'No' })

Write-Host ""
Write-Host ("{0,-$labelWidth}" -f "Source: ") -NoNewline -ForegroundColor DarkCyan
Write-Host $source -ForegroundColor White

Write-Host ("{0,-$labelWidth}" -f "Destination: ") -NoNewline -ForegroundColor DarkCyan
Write-Host $dest -ForegroundColor White

Write-Host ("{0,-$labelWidth}" -f "Delete sources:") -NoNewline -ForegroundColor DarkCyan
Write-Host (" " + $deleteText) -ForegroundColor White
Write-Host ""

$flags = @()
if ($yes) { $flags += '-yes' }
elseif ($no) { $flags += '-no' }
$quickCmd = ".\make-chd.ps1 -source `"$source`" -dest `"$dest`" " + ($flags -join ' ')

# Log file in the directory the script was executed from (current location)
$logStamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
$logPath = Join-Path (Get-Location).Path ("make-chd-log-{0}.log" -f $logStamp)

Write-Info "Log file: $logPath"
Write-Info "Found $($inputs.Count) compatible file(s)."
Write-Host ""

Write-Host "Quick Start Command:" -ForegroundColor DarkYellow
Write-Host $quickCmd.Trim() -ForegroundColor Cyan
Write-Host ""

$quickCmdLine = $quickCmd.Trim()
@(
  $quickCmdLine
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
      if (Test-ChdmanVerify -chdPath $outPath) {
        Write-Warn "Skipping (already valid): $outPath"
        Add-Content -LiteralPath $logPath -Value $inputPath -Encoding utf8
        Add-Content -LiteralPath $logPath -Value ("SKIPPED (already valid)    {0}" -f $outPath) -Encoding utf8
        Add-Content -LiteralPath $logPath -Value '' -Encoding utf8
        $skipped++
        continue
      } else {
        Write-Warn "Existing CHD failed verify; recreating: $outPath"
      }
    }

    Write-Info "Creating CHD from: $inputPath"
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

    Write-Summary "Verified: $outPath"
    Add-Content -LiteralPath $logPath -Value $inputPath -Encoding utf8
    Add-Content -LiteralPath $logPath -Value ("VERIFIED    {0}" -f $outPath) -Encoding utf8
    Add-Content -LiteralPath $logPath -Value '' -Encoding utf8
    $success++

    if ($deleteSources) {
      $toDelete = New-Object System.Collections.Generic.List[string]
      $toDelete.Add($inputPath) | Out-Null

      if ($ext -eq ".cue") {
        $bins = Get-CueReferencedBins -cuePath $inputPath
        foreach ($b in $bins) {
          $toDelete.Add($b) | Out-Null
        }
      }

      foreach ($p in ($toDelete | Sort-Object -Unique)) {
        if (Test-Path -LiteralPath $p -PathType Leaf) {
          Write-Info "Deleting source: $p"
          Remove-Item -LiteralPath $p -Force -ErrorAction Stop
        }
      }
    }
  } catch {
    $failed++
    Write-Host "Failed: $inputPath" -ForegroundColor Red
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

