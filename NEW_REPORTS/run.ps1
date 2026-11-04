# run.ps1 - Windows PowerShell launcher. Always writes both outputs:
#   reports\GT_Report_<from>_to_<to>.xlsx   full history (+ live TWS data)
#   reports\daily\GT_Report_<day>.xlsx      one workbook per trading day
# Extra arguments are passed to the script, e.g.
#   .\run.ps1 --from 2026-10-01 --to 2026-10-31 --no-tws

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$VenvDir = ".venv"
$Script  = "gt_flex_report.py"

# The script reads .env itself
if (-not (Test-Path ".env")) {
    Write-Error "Error: .env file not found. Copy .env.example to .env and fill in credentials."
    exit 1
}

if (-not (Test-Path $VenvDir)) {
    Write-Host "Creating virtual environment..."
    python -m venv $VenvDir
}

& "$VenvDir\Scripts\Activate.ps1"
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet

Write-Host "Running $Script..."
python $Script --daily @args
