param(
    [ValidateSet('start', 'stop', 'status', 'check', 'backup')]
    [string]$Action = 'start'
)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
Set-Location -LiteralPath $projectRoot
$localPython = $null
foreach ($relativePath in @('.local/venv/Scripts/python.exe', '.codex-venv/Scripts/python.exe', '.venv/Scripts/python.exe', 'venv/Scripts/python.exe')) {
    $candidate = Join-Path $projectRoot $relativePath
    if (Test-Path -LiteralPath $candidate) {
        & $candidate -c "import flask, psycopg, reportlab, openpyxl, googleapiclient" 2>$null
        if ($LASTEXITCODE -eq 0) { $localPython = $candidate; break }
    }
}
if (-not $localPython) {
    Write-Host 'Preparando Python y las dependencias locales...'
    $localVenv = Join-Path $projectRoot '.local/venv'
    python -m venv $localVenv
    if ($LASTEXITCODE -ne 0) { throw 'Se necesita Python 3.10 o superior para preparar SIG local.' }
    $localPython = Join-Path $localVenv 'Scripts/python.exe'
    & $localPython -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) { throw 'No se pudieron instalar las dependencias.' }
}
$env:PYTHONIOENCODING = 'utf-8'
& $localPython scripts/local.py $Action
if ($LASTEXITCODE -ne 0) { throw 'SIG local no pudo completar la operacion. Ver .local/web.log y .local/postgres.log.' }
