<#
    Starts the FlexKit grid harness from the folder this script sits in.
    Reads the exported JSON files under .\Data only: no database, no network
    calls, no credentials.

    Usage:  .\run-bench.ps1              (port 5399, opens the browser)
            .\run-bench.ps1 -Port 5400   (different port)
            .\run-bench.ps1 -NoBrowser   (do not launch a browser)
#>

[CmdletBinding()]
param(
    [int]$Port = 5399,
    [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'

$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$exe = Join-Path $root 'FlexKitTester.exe'
$dataDir = Join-Path $root 'Data'
$baseUrl = "http://127.0.0.1:$Port"
$itemsUrl = "$baseUrl/edit-items"
$modelUrl = "$baseUrl/edit-model-options"

function Stop-WithMessage([string]$Message) {
    Write-Host ''
    Write-Host $Message -ForegroundColor Red
    Write-Host ''
    Read-Host 'Press Enter to close'
    exit 1
}

Write-Host ''
Write-Host '=== FlexKit grid harness ===' -ForegroundColor Cyan
Write-Host "Folder: $root"

if (-not (Test-Path -LiteralPath $exe)) {
    Stop-WithMessage "FlexKitTester.exe is not in $root. Unzip the whole package and run this script from inside the unzipped folder."
}

foreach ($name in @('estimating-items-full.json', 'assembly-details-full.json')) {
    if (-not (Test-Path -LiteralPath (Join-Path $dataDir $name))) {
        Stop-WithMessage "Data\$name is missing. Without it the harness shows a red DATA NOT LOADED banner instead of real rows."
    }
}

# Files extracted from a downloaded ZIP carry the mark-of-the-web; without this
# Windows blocks the exe before it starts.
try {
    Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
} catch { }

$portInUse = $false
try {
    if (Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop) { $portInUse = $true }
} catch {
    $probe = New-Object System.Net.Sockets.TcpClient
    try { $probe.Connect('127.0.0.1', $Port); $portInUse = $true } catch { } finally { $probe.Dispose() }
}

if ($portInUse) {
    Stop-WithMessage "TCP port $Port is already in use on this machine. Close whatever is listening on port $Port, or start on another port:`r`n    .\run-bench.ps1 -Port 5400"
}

# Expected build marker, computed the same way both screens compute theirs.
# Read it back to whoever built the package: a mismatch means a stale build.
try {
    $ticks = [long]([System.IO.File]::GetLastWriteTimeUtc((Join-Path $root 'FlexCore.dll')).Ticks)
    $seconds = [long](($ticks - ($ticks % 10000000)) / 10000000)
    $names = @('Light Blue', 'Mint Green', 'Peach', 'Lavender', 'Lemon', 'Pink', 'Aqua', 'Sand')
    Write-Host ("Expected header: build {0} - selection: {1}" -f ($seconds % 100000), $names[[int]($seconds % 8)])
} catch { }

# Passed on the command line rather than through the environment so nothing
# depends on variable inheritance.
$appArgs = @('--urls', $baseUrl, '--environment', 'Production')

Write-Host ''
Write-Host 'Screens:' -ForegroundColor Cyan
Write-Host "  Edit Items DB           $itemsUrl"
Write-Host "  Edit Models and Options $modelUrl"
Write-Host ''
Write-Host 'Leave this window open. Press Ctrl+C here to stop the harness.' -ForegroundColor Yellow
Write-Host ''

$process = $null
try {
    # The working directory must be the app folder: ASP.NET Core takes its
    # content root from the process working directory.
    $process = Start-Process -FilePath $exe -ArgumentList $appArgs -WorkingDirectory $root -NoNewWindow -PassThru

    if (-not $NoBrowser) {
        for ($i = 0; $i -lt 60; $i++) {
            Start-Sleep -Milliseconds 500
            if ($process.HasExited) { break }
            try {
                Invoke-WebRequest -Uri $itemsUrl -UseBasicParsing -TimeoutSec 3 | Out-Null
                Start-Process $itemsUrl
                break
            } catch { }
        }
    }

    while (-not $process.HasExited) { Start-Sleep -Milliseconds 400 }
}
finally {
    if ($process) {
        if (-not $process.HasExited) {
            Write-Host ''
            Write-Host 'Stopping the harness...' -ForegroundColor Yellow
            if (-not $process.WaitForExit(5000)) {
                try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch { }
            }
        }
        Write-Host ''
        Write-Host ("Harness stopped (exit code {0})." -f $process.ExitCode)
    }
    Read-Host 'Press Enter to close'
}
