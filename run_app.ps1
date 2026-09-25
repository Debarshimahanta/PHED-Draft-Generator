$ErrorActionPreference = "Stop"

$ProjectRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location $ProjectRoot

$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

function Get-SystemPython {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return @("py", "-3")
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python -and $python.Source -notmatch "WindowsApps") {
        return @("python")
    }

    return $null
}

if (-not (Test-Path $VenvPython)) {
    $PythonCmd = Get-SystemPython

    if (-not $PythonCmd) {
        Write-Host ""
        Write-Host "Python is not installed or Windows is redirecting 'python' to the Microsoft Store." -ForegroundColor Red
        Write-Host ""
        Write-Host "Install Python 3.11 or 3.12 from https://www.python.org/downloads/windows/" -ForegroundColor Yellow
        Write-Host "IMPORTANT: Tick 'Add python.exe to PATH' during installation." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "After installation, close this PowerShell window, open a new one, and run this launcher again."
        Write-Host ""
        Read-Host "Press Enter to close"
        exit 1
    }

    Write-Host "Creating virtual environment..."
    if ($PythonCmd.Count -eq 2) {
        & $PythonCmd[0] $PythonCmd[1] -m venv .venv
    } else {
        & $PythonCmd[0] -m venv .venv
    }

    if (-not (Test-Path $VenvPython)) {
        throw "Virtual environment creation failed."
    }

    Write-Host "Installing dependencies..."
    & $VenvPython -m pip install --upgrade pip
    & $VenvPython -m pip install -r requirements.txt
}

Write-Host "Starting PHED Draft Generator..."
Write-Host "Open http://localhost:8501 in your browser if it does not open automatically."

Start-Process "http://localhost:8501"
& $VenvPython -m streamlit run app/main.py
