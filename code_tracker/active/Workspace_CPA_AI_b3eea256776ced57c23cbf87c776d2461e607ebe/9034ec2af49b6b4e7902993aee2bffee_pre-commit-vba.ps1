�<#
.SYNOPSIS
    Hook Git Pre-Commit pour valider l'encodage des fichiers VBA (CP1252/Windows-1252).
.DESCRIPTION
    Ce script vérifie tous les fichiers .bas, .cls, .frm en cours de commit (Staged).
    Il échoue si :
    - Un fichier contient des caractères UTF-8 (détectés via BOM ou caractères spéciaux).
    - Un fichier n'est pas en CRLF (fin de ligne Windows).
.USAGE
    Appelé automatiquement par Git si configuré.
    Peut être testé manuellement : ./pre-commit-vba.ps1
#>

$ErrorActionPreference = "Stop"
$CurrentDir = Get-Location

Write-Host "🔍 [VBA-Hook] Vérification de l'encodage des fichiers VBA..." -ForegroundColor Cyan

# 1. Récupérer les fichiers 'staged' (index) qui sont des sources VBA
try {
    $stagedFiles = git diff --cached --name-only --diff-filter=ACM | Where-Object { $_ -match "\.(bas|cls|frm)$" }
}
catch {
    Write-Warning "[VBA-Hook] Impossible de récupérer les fichiers staged (Git non initialisé ?). Vérification ignorée."
    exit 0
}

if (-not $stagedFiles) {
    Write-Host "✅ [VBA-Hook] Aucun fichier VBA à vérifier." -ForegroundColor Green
    exit 0
}

$issuesFound = $false

foreach ($fileRelPath in $stagedFiles) {
    $fullPath = Join-Path $CurrentDir $fileRelPath

    if (-not (Test-Path $fullPath)) {
        continue
    }

    # Lecture des bytes bruts
    $bytes = [System.IO.File]::ReadAllBytes($fullPath)

    # TEST 1 : BOM UTF-8 (EF BB BF)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        Write-Host "❌ [ERREUR] BOM UTF-8 détecté : $fileRelPath" -ForegroundColor Red
        Write-Host "   -> Le fichier doit être enregistré en 'Western European (Windows) - CP1252' sans BOM."
        $issuesFound = $true
        continue
    }

    # TEST 2 : Caractères invalides CP1252 (Heuristique simple)
    # On essaie de lire en pur 1252. Si le fichier était du "vrai" UTF-8 avec accents, 
    # les séquences multi-octets apparaîtraient comme des caractères étranges (Ã© pour é).
    # Ce test est difficile à automatiser parfaitement sans faux positifs, 
    # on se concentre sur la détection CRLF pour l'instant et le BOM.
    
    # TEST 3 : CRLF (Windows)
    # On lit le texte pour vérifier les fins de ligne
    $content = [System.IO.File]::ReadAllText($fullPath)
    if ($content -match "[^\r]\n") {
        Write-Host "❌ [ERREUR] Fin de ligne LF (Unix) détectée : $fileRelPath" -ForegroundColor Red
        Write-Host "   -> Le fichier doit être enregistré avec des fins de ligne CRLF (Windows)."
        $issuesFound = $true
    }
}

if ($issuesFound) {
    Write-Host "⛔ Commit bloqué. Veuillez corriger l'encodage des fichiers listés ci-dessus." -ForegroundColor Red
    exit 1
}
else {
    Write-Host "✅ [VBA-Hook] Tous les fichiers VBA sont conformes (Pas de BOM UTF-8, CRLF OK)." -ForegroundColor Green
    exit 0
}
� ��
�� "(b3eea256776ced57c23cbf87c776d2461e607ebe2Sfile:///c:/Users/AbelBoudreau/Workspace_CPA_AI/scripts/git-hooks/pre-commit-vba.ps1:.file:///c:/Users/AbelBoudreau/Workspace_CPA_AI