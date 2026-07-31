<#
.SYNOPSIS
    Launches the SSAS -> Fabric Migration Web UI with an elevated (admin)
    token, automatically.

.DESCRIPTION
    Analysis Services (AMO) calls made by Phase 1 steps (e.g. Extract)
    require the connecting Windows account to be recognized as an SSAS
    Server Administrator. On many machines, membership in the local
    Administrators group is only active in an *elevated* process (User
    Account Control filters it out of a normal/non-elevated session) -
    and on machines that are not joined to the account's AD domain (so
    SQL Server Management Studio's account picker cannot validate/add
    the specific domain account as a named admin), running elevated is
    the only way to satisfy this requirement.

    This script checks whether it is already running elevated; if not,
    it re-launches itself via UAC (one confirmation prompt) and exits the
    original, non-elevated instance. Once elevated, it starts the
    Streamlit Web UI. Every pipeline step subsequently run by clicking a
    button in the browser is a child process of this elevated session, so
    it inherits the same admin-capable token - no separate elevated
    terminal, SSMS, or manual step outside the UI is required.

.NOTES
    Run this from the repo root (double-click run_ui.bat, or run
    `.\run_ui.ps1` from a PowerShell prompt already in this folder).
#>

param(
    [string]$VenvUiPython = ".venv-ui\Scripts\python.exe",
    [string]$AppPath = "ssas_fabric_migrator\ui\app.py"
)

$ErrorActionPreference = "Stop"
$repoRoot = $PSScriptRoot

function Test-IsElevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsElevated)) {
    Write-Host "Requesting elevation (one-time Windows UAC prompt) so SSAS/AMO steps run with admin rights..." -ForegroundColor Yellow
    $psExePath = (Get-Process -Id $PID).Path
    $argList = @(
        "-NoExit",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`"",
        "-VenvUiPython", "`"$VenvUiPython`"",
        "-AppPath", "`"$AppPath`""
    )
    Start-Process -FilePath $psExePath -Verb RunAs -ArgumentList $argList
    exit
}

Set-Location $repoRoot

$pythonExe = Join-Path $repoRoot $VenvUiPython
$appFile = Join-Path $repoRoot $AppPath

if (-not (Test-Path $pythonExe)) {
    Write-Error "Could not find '$pythonExe'. Create the .venv-ui environment first (see README Quickstart step 4)."
    exit 1
}
if (-not (Test-Path $appFile)) {
    Write-Error "Could not find '$appFile'. Make sure run_ui.ps1 lives in the repo root alongside ssas_fabric_migrator\."
    exit 1
}

Write-Host "Running elevated. Launching Web UI from $repoRoot ..." -ForegroundColor Green
& $pythonExe -m streamlit run $appFile
