$ErrorActionPreference = "Stop"

$ProjectRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
Set-Location $ProjectRoot

$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

function Get-SystemPython {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { return @("py", "-3") }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python -and $python.Source -notmatch "WindowsApps") { return @("python") }

    return $null
}

if (-not (Test-Path $VenvPython)) {
    $PythonCmd = Get-SystemPython

    if (-not $PythonCmd) {
        Write-Host ""
        Write-Host "Python is not available to PowerShell." -ForegroundColor Red
        Write-Host "Please install Python 3.11/3.12 or add it to PATH, then reopen PowerShell." -ForegroundColor Yellow
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

$StreamlitArgs = @(
    "-m", "streamlit", "run", "app/main.py",
    "--server.address", "127.0.0.1",
    "--server.port", "8501",
    "--browser.gatherUsageStats", "false"
)

$proc = Start-Process -FilePath $VenvPython -ArgumentList $StreamlitArgs -PassThru -NoNewWindow

$Ready = $false
for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Seconds 1

    if ($proc.HasExited) {
        Write-Host ""
        Write-Host "The app stopped before starting. Review the error messages above." -ForegroundColor Red
        Read-Host "Press Enter to close"
        exit 1
    }

    try {
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:8501/_stcore/health" -UseBasicParsing -TimeoutSec 2
        if ($response.StatusCode -eq 200) {
            $Ready = $true
            break
        }
    } catch {
        # Keep waiting while Streamlit starts.
    }
}

if ($Ready) {
    Write-Host "PHED Draft Generator is ready." -ForegroundColor Green
    Write-Host "Opening http://localhost:8501"
    Start-Process "http://localhost:8501"
} else {
    Write-Host ""
    Write-Host "The app did not become ready within 60 seconds." -ForegroundColor Red
    Write-Host "Please review the PowerShell output above for the actual error."
}

Write-Host ""
Write-Host "Keep this PowerShell window open while using the app."
Write-Host "Press Ctrl+C here when you want to stop it."

Wait-Process -Id $proc.Id
