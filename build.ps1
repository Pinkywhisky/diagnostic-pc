# =========================
# BUILD DEBUG_PC.EXE
# =========================

$ErrorActionPreference = "Stop"

try {
    Write-Host ""
    Write-Host "=========================" -ForegroundColor Cyan
    Write-Host "   BUILD DEBUG_PC.EXE    " -ForegroundColor Cyan
    Write-Host "=========================" -ForegroundColor Cyan
    Write-Host ""

    # Chemins
    $scriptPath = Join-Path $PSScriptRoot "debug_pc.ps1"
    $exePath    = Join-Path $PSScriptRoot "debug_pc.exe"

    # Vérif script
    if (-not (Test-Path $scriptPath)) {
        throw "debug_pc.ps1 introuvable dans $PSScriptRoot"
    }

    # Module ps2exe
    if (-not (Get-Module -ListAvailable -Name ps2exe)) {
        Write-Host "[INFO] Installation du module ps2exe..." -ForegroundColor Yellow
        Install-Module ps2exe -Scope CurrentUser -Force
    }

    Write-Host "[INFO] Chargement du module..." -ForegroundColor Yellow
    Import-Module ps2exe -Force

    # Suppression ancien exe
    if (Test-Path $exePath) {
        Write-Host "[INFO] Suppression ancien EXE..." -ForegroundColor Yellow
        Remove-Item $exePath -Force
    }

    # Build
    Write-Host "[INFO] Compilation en cours..." -ForegroundColor Cyan

    Invoke-ps2exe `
        -inputFile $scriptPath `
        -outputFile $exePath `
        -noConsole `
        -title "Diagnostic PC" `
        -description "Outil de diagnostic poste" `
        -company "Ecritel"

    # Vérif résultat
    if (Test-Path $exePath) {
        Write-Host ""
        Write-Host "[SUCCESS] BUILD OK" -ForegroundColor Green
        Write-Host $exePath -ForegroundColor Gray
    }
    else {
        throw "Le fichier EXE n'a pas été généré"
    }

}
catch {
    Write-Host ""
    Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    pause
    exit 1
}