<#
.SYNOPSIS
Runs the Log Report Generator in headless automated batch mode.

.DESCRIPTION
You can right-click and "Run with PowerShell", or call this script from another process.
It accepts a path to a logs directory as its first argument.
#>

param(
    [Parameter(Position=0)]
    [string]$LogsDir
)

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "    Log Report Generator - Automated Batch Mode" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -Path $ScriptDir

# Resolve python executable natively
$PythonExe = "python"
if (Test-Path "portable_python\python.exe") {
    $PythonExe = "portable_python\python.exe"
    Write-Host "[INFO] Using packaged Portable Python engine." -ForegroundColor DarkGray
} elseif (Test-Path "..\.venv\Scripts\python.exe") {
    $PythonExe = "..\.venv\Scripts\python.exe"
    Write-Host "[INFO] Using virtual environment python." -ForegroundColor DarkGray
} elseif (Test-Path ".venv\Scripts\python.exe") {
    $PythonExe = ".venv\Scripts\python.exe"
    Write-Host "[INFO] Using virtual environment python." -ForegroundColor DarkGray
}

if ($LogsDir) {
    Write-Host "[INFO] Detected target Log Folder: $LogsDir" -ForegroundColor Blue
    & $PythonExe batch_process.py --logs-dir "$LogsDir"
} else {
    Write-Host "[INFO] Running default software logs scan mode..." -ForegroundColor Blue
    & $PythonExe batch_process.py
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[ERROR] Report generation encountered an issue." -ForegroundColor Red
} else {
    Write-Host "`n[SUCCESS] Report generation complete." -ForegroundColor Green
}

Write-Host ""
Write-Host "Press any key to exit..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
