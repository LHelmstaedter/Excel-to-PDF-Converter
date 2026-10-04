# Builds a single windowed EXE (no console window): dist\ExcelPdfExporter.exe
# Usage: .\build.ps1            (uses "python" from PATH)
#        .\build.ps1 -Python "C:\path\to\python.exe"
param([string]$Python = "python")

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

& $Python -m pip install -r requirements.txt pyinstaller
if ($LASTEXITCODE -ne 0) { throw "pip install failed" }

& $Python -m PyInstaller --noconfirm --clean --windowed --onefile `
    --name ExcelPdfExporter `
    --collect-all tkinterdnd2 `
    --specpath build `
    app.pyw
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

Write-Host "Done: dist\ExcelPdfExporter.exe"
