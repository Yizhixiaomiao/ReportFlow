param(
    [switch]$InstallDeps
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$candidates = @(
    @("$PSScriptRoot\.venv\Scripts\python.exe"),
    @("py", "-3.11"),
    @("$PSScriptRoot\venv\Scripts\python.exe"),
    @("python")
)

$python = $null
foreach ($candidate in $candidates) {
    try {
        $exe = $candidate[0]
        $args = @()
        if ($candidate.Count -gt 1) {
            $args = $candidate[1..($candidate.Count - 1)]
        }
        & $exe @args -c "import sys; print(sys.executable)" *> $null
        if ($LASTEXITCODE -eq 0) {
            $python = $candidate
            break
        }
    } catch {
        continue
    }
}

if (-not $python) {
    throw "No usable Python found. Install Python 3.11 or create desktop\.venv first."
}

function Invoke-SelectedPython {
    param([string[]]$PythonCommand, [string[]]$PythonArgs)
    $exe = $PythonCommand[0]
    $prefixArgs = @()
    if ($PythonCommand.Count -gt 1) {
        $prefixArgs = $PythonCommand[1..($PythonCommand.Count - 1)]
    }
    & $exe @prefixArgs @PythonArgs
}

Write-Host "Using Python:"
Invoke-SelectedPython $python @("-c", "import sys; print(sys.version); print(sys.executable)")

$legacySitePackages = "$PSScriptRoot\.venv\Lib\site-packages"
if (Test-Path $legacySitePackages) {
    if ($env:PYTHONPATH) {
        $env:PYTHONPATH = "$legacySitePackages;$env:PYTHONPATH"
    } else {
        $env:PYTHONPATH = $legacySitePackages
    }
}

if ($InstallDeps) {
    Invoke-SelectedPython $python @("-m", "pip", "install", "-r", "requirements.txt")
}

$dependencyOk = $true
try {
    Invoke-SelectedPython $python @("-c", "import openpyxl, win32com.client, pythoncom, PyInstaller") *> $null
    if ($LASTEXITCODE -ne 0) {
        $dependencyOk = $false
    }
} catch {
    $dependencyOk = $false
}

if (-not $dependencyOk) {
    throw "The selected Python is missing dependencies. Run: .\build_exe.ps1 -InstallDeps"
}

Invoke-SelectedPython $python @("-m", "PyInstaller", "--noconfirm", "--clean", "--windowed", "--name", "ReportFlowDesktop", "reportflow_desktop.py")
Write-Host "Build complete: $PSScriptRoot\dist\ReportFlowDesktop\ReportFlowDesktop.exe"
