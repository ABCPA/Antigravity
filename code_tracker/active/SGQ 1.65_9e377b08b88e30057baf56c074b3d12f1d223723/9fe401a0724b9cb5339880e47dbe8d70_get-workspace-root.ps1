�<#
.SYNOPSIS
Détecte automatiquement la racine du workspace SGQ

.DESCRIPTION
Recherche la racine du projet en remontant l'arborescence jusqu'à trouver
les marqueurs du projet (sgq-config.json, VERSION.txt, index.xlsm).
Compatible avec structure actuelle (SGQ 1.65) et futures (SGQ, SGQ/dev).

.PARAMETER StartPath
Chemin de départ (défaut: répertoire courant)

.OUTPUTS
String - Chemin absolu vers la racine du workspace

.EXAMPLE
$root = & ".\get-workspace-root.ps1"
Write-Host "Racine: $root"

.EXAMPLE
$root = & ".\get-workspace-root.ps1" -StartPath "C:\VBA\SGQ 1.65\scripts\migration"
# Retourne: C:\VBA\SGQ 1.65
#>

param(
    [string]$StartPath = $PWD.Path
)

function Find-WorkspaceRoot {
    param([string]$Path)

    $currentPath = $Path
    $maxDepth = 10
    $depth = 0

    while ($depth -lt $maxDepth) {
        # Vérifier marqueurs du projet
        $markers = @(
            "sgq-config.json",
            "VERSION.txt",
            "index.xlsm",
            "vba-files"
        )

        $foundMarkers = 0
        foreach ($marker in $markers) {
            $markerPath = Join-Path $currentPath $marker
            if (Test-Path $markerPath) {
                $foundMarkers++
            }
        }

        # Si au moins 2 marqueurs trouvés, c'est la racine
        if ($foundMarkers -ge 2) {
            return $currentPath
        }

        # Remonter d'un niveau
        $parentPath = Split-Path $currentPath -Parent

        # Si on atteint la racine du disque, arrêter
        if (-not $parentPath -or $parentPath -eq $currentPath) {
            break
        }

        $currentPath = $parentPath
        $depth++
    }

    # Fallback: retourner le chemin de départ
    Write-Warning "Impossible de détecter la racine du workspace. Utilisation de: $Path"
    return $Path
}

# Exécution
$workspaceRoot = Find-WorkspaceRoot -Path $StartPath

# Normaliser le chemin (résoudre .., symlinks, etc.)
$workspaceRoot = (Resolve-Path $workspaceRoot).Path

# Retourner le résultat
return $workspaceRoot
�*cascade08"(9e377b08b88e30057baf56c074b3d12f1d2237232?file:///c:/VBA/SGQ%201.65/scripts/common/get-workspace-root.ps1:file:///c:/VBA/SGQ%201.65