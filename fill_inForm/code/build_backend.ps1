# Fill Form backend packaging script for Windows.
# Run it from PowerShell or via the accompanying batch launcher:
#   .\build_backend.ps1
# PaddleOCR runtime is bundled, but downloaded OCR models are not.

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$pythonCandidates = @(
    (Join-Path $projectRoot '.venv\Scripts\python.exe'),
    (Join-Path $scriptRoot '.venv\Scripts\python.exe')
)
$pythonPath = $pythonCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

if (-not $pythonPath) {
    Write-Host 'Project Python environment was not found in either of these locations:' -ForegroundColor Red
    $pythonCandidates | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    Write-Host 'Check the PyCharm interpreter or create the .venv environment first.' -ForegroundColor Yellow
    Read-Host 'Press Enter to exit'
    exit 1
}

Write-Host "Using Python: $pythonPath" -ForegroundColor Cyan
& $pythonPath --version
if ($LASTEXITCODE -ne 0) {
    Write-Host 'The virtual environment Python cannot start.' -ForegroundColor Red
    Write-Host 'In PyCharm, set the project interpreter to a valid Python installation, then recreate .venv if needed.' -ForegroundColor Yellow
    Read-Host 'Press Enter to exit'
    exit 1
}

Write-Host 'Checking PyInstaller...' -ForegroundColor Cyan
# Python returns a non-zero exit code while probing a missing module. Temporarily
# allow that expected result so the script can install PyInstaller below.
$savedErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
& $pythonPath -m PyInstaller --version 2>$null
$pyInstallerInstalled = $LASTEXITCODE -eq 0
$ErrorActionPreference = $savedErrorActionPreference
if (-not $pyInstallerInstalled) {
    Write-Host 'Installing PyInstaller...' -ForegroundColor Cyan
    & $pythonPath -m pip install PyInstaller
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller installation failed. Check the Python environment and network connection.' }
}

Set-Location -LiteralPath $scriptRoot
Write-Host 'Building backend. Downloaded OCR models are not included...' -ForegroundColor Cyan
& $pythonPath -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --name fill_form_backend `
    --collect-all paddle `
    --collect-all paddleocr `
    --collect-all paddlex `
    --collect-all fitz `
    --collect-all aistudio_sdk `
    --collect-all chardet `
    --collect-all colorlog `
    --collect-all filelock `
    --collect-all huggingface_hub `
    --collect-all modelscope `
    --collect-all pandas `
    --collect-all prettytable `
    --collect-all cpuinfo `
    --collect-all pydantic `
    --collect-all yaml `
    --collect-all requests `
    --collect-all ruamel `
    --collect-all typing_extensions `
    --collect-all ujson `
    --collect-all imagesize `
    --collect-all cv2 `
    --collect-all pyclipper `
    --collect-all pypdfium2 `
    --collect-all bidi `
    --collect-all shapely `
    --hidden-import uvicorn.logging `
    --hidden-import uvicorn.loops.auto `
    --hidden-import uvicorn.protocols.http.auto `
    --hidden-import uvicorn.protocols.websockets.auto `
    main.py

if ($LASTEXITCODE -ne 0) { throw 'Backend packaging failed. See the error details above.' }

$outputDirectory = Join-Path $scriptRoot 'dist\fill_form_backend'
Write-Host ''
Write-Host 'Packaging complete. Backend directory:' -ForegroundColor Green
Write-Host $outputDirectory -ForegroundColor Green
Write-Host 'This directory will be included as an Electron application resource.' -ForegroundColor Yellow
Read-Host 'Press Enter to exit'
