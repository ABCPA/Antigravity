�<#
.SYNOPSIS
    Connecte le Workspace CPA AI aux sources de données réelles (Z:, SharePoint, Admin).
    Utilise des Raccourcis (.lnk) au lieu de Symlinks pour éviter les problèmes de permissions Admin.

.DESCRIPTION
    Structure créée :
    Workspace_CPA_AI/
      ├── 1_Clients/               ---> (Shortcut) Z:\CPA
      ├── 2_Normes_References/     ---> (Shortcut) ...\Doc utiles
      └── 3_Modeles_Livrables/     ---> (Shortcut) ...\Administration\caseware & Modeles

.NOTES
    Exécutez ce script pour initialiser l'environnement "Live Data".
#>

$WorkspaceRoot = "C:\Users\AbelBoudreau\Workspace_CPA_AI"

# Configuration des Mappings
$Mappings = @{
    "1_Clients"                         = "Z:\"
    "2_Normes_References"               = "C:\Users\AbelBoudreau\Abel Boudreau\intranet - Documents\Doc utiles"
    "3_Modeles_Livrables\Partage_Admin" = "C:\Users\AbelBoudreau\Abel Boudreau\Administration - Documents"
}

# Fonction helper pour créer un raccourci
function New-Shortcut {
    param (
        [string]$SourcePath,
        [string]$DestinationName, # Nom du raccourci (sans .lnk)
        [string]$ParentDir
    )

    if (-not (Test-Path $SourcePath)) {
        Write-Warning "Source introuvable : $SourcePath. Raccourci ignoré."
        return
    }

    if (-not (Test-Path $ParentDir)) {
        New-Item -Path $ParentDir -ItemType Directory -Force | Out-Null
    }

    $WshShell = New-Object -ComObject WScript.Shell
    $ShortcutPath = Join-Path $ParentDir "$DestinationName.lnk"
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $SourcePath
    $Shortcut.Description = "Lien vers $SourcePath (Auto-généré par CPA AI)"
    $Shortcut.Save()

    Write-Host "  [+] Connecté : $DestinationName -> $SourcePath" -ForegroundColor Green
}

Write-Host "Initialisation de la Data Fabric CPA AI..." -ForegroundColor Cyan

# 1. Clients
New-Shortcut -SourcePath $Mappings["1_Clients"] -DestinationName "ACCES_DISQUE_CLIENTS_Z" -ParentDir "$WorkspaceRoot\1_Clients"

# 2. Normes (Intranet)
New-Shortcut -SourcePath $Mappings["2_Normes_References"] -DestinationName "ACCES_INTRANET_DOCS" -ParentDir "$WorkspaceRoot\2_Normes_References"

# 3. Modèles (Admin)
New-Shortcut -SourcePath $Mappings["3_Modeles_Livrables\Partage_Admin"] -DestinationName "ACCES_ADMIN_DOCS" -ParentDir "$WorkspaceRoot\3_Modeles_Livrables"

Write-Host "`nArchitecture connectée avec succès." -ForegroundColor Cyan
Write-Host "Les agents peuvent maintenant suivre ces raccourcis pour indexer ou consulter les données." -ForegroundColor Gray
�*cascade08"(40ee8fe1244e0f95119fc9e995da074dedc0b1c12Nfile:///C:/Users/AbelBoudreau/Workspace_CPA_AI/scripts/Connect-DataSources.ps1:.file:///C:/Users/AbelBoudreau/Workspace_CPA_AI