�# Script pour ajouter l'attribut VBA pour cacher les macros dans ALT+F8
# Objectif: Cacher toutes les procédures Public Sub qui ne commencent pas par "btn"

$ErrorActionPreference = "Stop"
$modulesPath = "C:\VBA\SGQ 1.65.worktrees\worktree-2025-12-16T20-14-20\vba-files\Module"

# Liste des fichiers à traiter
$filesToProcess = @(
    "modWorkbookHandlers.bas",
    "modSGQUtilitaires.bas",
    "modSGQCreation.bas",
    "modSGQProtection.bas",
    "modExcelUtils.bas",
    "modDiagnostics.bas",
    "modAppStateGuard.bas",
    "modSGQContextuel.bas",
    "modSGQChangeTracking.bas",
    "modSGQAuditLog.bas",
    "modRibbonSGQ.bas",
    "modSGQUpdateManager.bas",
    "modSGQTrackingBuilder.bas",
    "modSGQRefactor.bas",
    "modSGQVBProjectHelpers.bas",
    "modSGQValidations.bas",
    "modLegacyAutoRefactor.bas",
    "modVBAInspector.bas",
    "modTestWorkbookEvents.bas",
    "modSGQUIActionDispatcher.bas"
)

$processedCount = 0
$skippedCount = 0

foreach ($file in $filesToProcess) {
    $filePath = Join-Path $modulesPath $file

    if (-not (Test-Path $filePath)) {
        Write-Warning "Fichier non trouvé: $filePath"
        $skippedCount++
        continue
    }

    Write-Host "Traitement: $file" -ForegroundColor Cyan

    # Lire le contenu du fichier
    $content = Get-Content $filePath -Raw -Encoding Default

    # Pattern pour trouver les Public Sub qui ne commencent PAS par btn
    # et qui n'ont pas déjà l'attribut
    $pattern = '(?m)^Public Sub (?!btn)([a-zA-Z_][a-zA-Z0-9_]*)\((.*?)\)\s*\r?\n(?!Attribute)'

    if ($content -match $pattern) {
        # Remplacer chaque occurrence
        $newContent = $content -replace $pattern, {
            param($match)
            $procName = $match.Groups[1].Value
            $params = $match.Groups[2].Value
            "Public Sub $procName($params)`r`nAttribute $procName.VB_ProcData.VB_Invoke_Func = "" \n14""`r`n"
        }

        # Sauvegarder
        Set-Content -Path $filePath -Value $newContent -Encoding Default -NoNewline
        Write-Host "  ✓ Modifié" -ForegroundColor Green
        $processedCount++
    } else {
        Write-Host "  - Aucune procédure publique trouvée (ou déjà modifié)" -ForegroundColor Yellow
        $skippedCount++
    }
}

Write-Host "`nRésumé:" -ForegroundColor White
Write-Host "  Fichiers modifiés: $processedCount" -ForegroundColor Green
Write-Host "  Fichiers ignorés: $skippedCount" -ForegroundColor Yellow
�*cascade08"(9e377b08b88e30057baf56c074b3d12f1d2237232?file:///c:/VBA/SGQ%201.65/scripts/add-hide-macro-attributes.ps1:file:///c:/VBA/SGQ%201.65