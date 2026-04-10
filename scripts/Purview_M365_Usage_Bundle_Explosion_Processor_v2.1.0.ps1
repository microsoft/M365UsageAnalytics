<#
.SYNOPSIS
    PowerShell launcher for Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.py

.DESCRIPTION
    Wrapper script that accepts PowerShell-style single-hyphen parameters, translates
    them to Python double-hyphen equivalents, and forwards them to the Python processor.
    Ensures Python 3.9+ is available and installs orjson if missing.

    Parameter name mapping (PowerShell → Python):
      -InputPath (-input, -i)    → --input
      -OutputDir (-output, -o)   → --output-dir
      -Mode (-m)                 → --mode
      -PromptFilter              → --prompt-filter
      -Reconcile                 → --reconcile
      -NoUserStats               → --no-userstats
      -quiet (-q)                → --quiet
      -version                   → --version

.PARAMETER InputPath
    (Required) Path to the input Purview audit log CSV file (must contain AuditData column).
    Aliases: -input, -i

.PARAMETER OutputDir
    Directory for output files. Default: same directory as the input file.
    Aliases: -output, -o

.PARAMETER Mode
    Processing mode: rollup (default) or event-level.
    Alias: -m

.PARAMETER PromptFilter
    Filter Copilot messages: Prompt | Response | Both | Null
    Maps to Python's --prompt-filter.

.PARAMETER Reconcile
    Run sample-based reconciliation after processing to validate rollup correctness.

.PARAMETER NoUserStats
    Skip generating UserStats and SessionCohort files (rollup mode only).

.PARAMETER quiet
    Suppress progress output (only errors are printed).
    Alias: -q

.PARAMETER version
    Show version and exit.

.EXAMPLE
    .\Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.ps1 -input Purview_Export.csv

.EXAMPLE
    .\Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.ps1 -i Purview_Export.csv -o ./output

.EXAMPLE
    .\Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.ps1 -input Purview_Export.csv -PromptFilter Prompt -quiet

.EXAMPLE
    .\Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.ps1 -input Purview_Export.csv -NoUserStats

.NOTES
    Version: 2.1.0
    Requires: PowerShell 7+ (pwsh), Python 3.9+ on PATH
    Optional: pip install orjson (5-10x faster JSON parsing)

    IMPORTANT: This wrapper requires PowerShell 7+ (pwsh.exe). It is not compatible
    with Windows PowerShell 5.1 (powershell.exe) due to encoding limitations.
#>

[CmdletBinding()]
param(
    [Alias('input','i')]
    [string]$InputPath,

    [Alias('output','o')]
    [string]$OutputDir,

    [Alias('m')]
    [ValidateSet('rollup', 'event-level')]
    [string]$Mode,

    [ValidateSet('Prompt', 'Response', 'Both', 'Null')]
    [string]$PromptFilter,

    [switch]$Reconcile,

    [switch]$NoUserStats,

    [Alias('q')]
    [switch]$quiet,

    [switch]$version
)

# ── Locate the Python script (same directory as this launcher) ───────────
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$PythonScript = Join-Path $ScriptDir 'Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.py'

if (-not (Test-Path $PythonScript)) {
    Write-Host "ERROR: Python script not found at: $PythonScript" -ForegroundColor Red
    exit 1
}

# ── Find Python ──────────────────────────────────────────────────────────
$PythonExe = $null
foreach ($candidate in @('python3', 'python', 'py')) {
    try {
        $found = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($found) {
            # Verify it's Python 3.9+
            $verOutput = & $found.Source --version 2>&1
            if ($verOutput -match 'Python (\d+)\.(\d+)') {
                $major = [int]$Matches[1]
                $minor = [int]$Matches[2]
                if ($major -ge 3 -and $minor -ge 9) {
                    $PythonExe = $found.Source
                    break
                }
            }
        }
    } catch {}
}

if (-not $PythonExe) {
    Write-Host "ERROR: Python 3.9+ not found on PATH. Install from https://python.org" -ForegroundColor Red
    Write-Host "       Ensure 'python' or 'python3' is available in your PATH." -ForegroundColor Yellow
    exit 1
}

# ── Ensure orjson is installed (5-10x faster JSON parsing) ───────────────
$orjsonCheck = & $PythonExe -c "import orjson" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "orjson not found — installing for faster JSON processing..." -ForegroundColor Yellow
    & $PythonExe -m pip install orjson --quiet 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  orjson installed successfully." -ForegroundColor Green
    } else {
        Write-Host "  orjson installation failed — the script will fall back to Python's built-in JSON parser (slower but functional)." -ForegroundColor Yellow
    }
}

# ── Build Python arguments from declared parameters ──────────────────────
$pyArgs = @($PythonScript)

if ($version) {
    $pyArgs += '--version'
    & $PythonExe @pyArgs
    exit $LASTEXITCODE
}

# Validate required -input parameter (not enforced via Mandatory so -version can run without it)
if (-not $PSBoundParameters.ContainsKey('InputPath') -or [string]::IsNullOrWhiteSpace($InputPath)) {
    Write-Host "ERROR: -input is required. Provide the path to the Purview audit log CSV." -ForegroundColor Red
    exit 1
}

# Required
$pyArgs += '--input'
$pyArgs += $InputPath

if ($PSBoundParameters.ContainsKey('OutputDir')) {
    $pyArgs += '--output-dir'
    $pyArgs += $OutputDir
}
if ($PSBoundParameters.ContainsKey('Mode')) {
    $pyArgs += '--mode'
    $pyArgs += $Mode
}
if ($PSBoundParameters.ContainsKey('PromptFilter')) {
    $pyArgs += '--prompt-filter'
    $pyArgs += $PromptFilter
}
if ($Reconcile) {
    $pyArgs += '--reconcile'
}
if ($NoUserStats) {
    $pyArgs += '--no-userstats'
}
if ($quiet) {
    $pyArgs += '--quiet'
}

Write-Host "Launching Python processor..." -ForegroundColor Cyan
Write-Host "  Python:  $PythonExe" -ForegroundColor DarkGray
Write-Host "  Script:  $PythonScript" -ForegroundColor DarkGray
Write-Host ""

& $PythonExe @pyArgs
exit $LASTEXITCODE
