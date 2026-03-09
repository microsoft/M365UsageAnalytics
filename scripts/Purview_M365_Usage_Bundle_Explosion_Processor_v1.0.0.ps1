<#
.SYNOPSIS
    PowerShell launcher for Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.py

.DESCRIPTION
    Wrapper script that passes all switches through to the Python explosion processor.
    Ensures Python is available and forwards arguments with identical names and formats.

.PARAMETER input
    Path to the input Purview audit log CSV file (must contain AuditData column).

.PARAMETER output
    Path for the output exploded CSV. Default: <input_stem>_Exploded.csv

.PARAMETER prompt-filter
    Filter Copilot messages: Prompt | Response | Both | Null

.PARAMETER workers
    Number of parallel workers. Default: min(cpu_count, 8). Use 1 to disable parallelism.

.PARAMETER chunk-size
    Number of CSV rows per processing chunk. Default: 5000.

.PARAMETER quiet
    Suppress progress output (only errors are printed).

.PARAMETER version
    Show version and exit.

.EXAMPLE
    .\Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.ps1 --input Purview_Export.csv
    .\Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.ps1 -i Purview_Export.csv -o Exploded.csv --workers 4
    .\Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.ps1 --input Purview_Export.csv --prompt-filter Prompt --workers 4 --quiet

.NOTES
    Version: 1.0.0
    Requires: Python 3.9+ on PATH
    Optional: pip install orjson (5-10x faster JSON parsing)
#>

[CmdletBinding()]
param()

# ── Locate the Python script (same directory as this launcher) ───────────
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$PythonScript = Join-Path $ScriptDir 'Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.py'

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

# ── Forward all arguments to Python ──────────────────────────────────────
# Pass through raw CLI arguments exactly as received (preserves --input, -i, --workers, etc.)
$pyArgs = @($PythonScript) + $args

Write-Host "Launching Python explosion processor..." -ForegroundColor Cyan
Write-Host "  Python:  $PythonExe" -ForegroundColor DarkGray
Write-Host "  Script:  $PythonScript" -ForegroundColor DarkGray
Write-Host ""

& $PythonExe @pyArgs
exit $LASTEXITCODE
