$ErrorActionPreference = "Stop"

# Works both when launched as a .ps1 file and when pasted into PowerShell.
$ProjectRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location $ProjectRoot

$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $VenvPython)) {
    Write-Host "Creating virtual environment..."
    python -m venv .venv

    Write-Host "Installing dependencies..."
    & $VenvPython -m pip install --upgrade pip
    & $VenvPython -m pip install -r requirements.txt
}

Write-Host "Starting PHED Draft Generator..."
Write-Host "Open http://localhost:8501 in your browser if it does not open automatically."

Start-Process "http://localhost:8501"
& $VenvPython -m streamlit run app/main.py
