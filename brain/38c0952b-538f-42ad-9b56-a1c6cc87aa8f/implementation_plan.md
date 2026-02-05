# Plan : Implémentation de Generate-AI-Context.ps1

Ce plan détaille la création d'un script PowerShell pour consolider les connaissances du projet dans un format optimisé pour NotebookLM.

## Proposed Changes

### [Scripts]

#### [NEW] [Generate-AI-Context.ps1](file:///c:/VBA/SGQ%201.65/scripts/Generate-AI-Context.ps1)
- Créer un script qui parcourt `vba-files/` et regroupe le code et la documentation.
- Inclure les fichiers `.bas`, `.cls`, `.md` (architecture, changelog).
- Sortie : `dist/Project-Context.md`.

## Verification Plan

### Automated Tests
- Exécuter le script et vérifier que le fichier de sortie n'est pas vide.
- Vérifier que la taille du fichier est raisonnable pour l'upload (NotebookLM a des limites par document).

### Manual Verification
- Inspecter le contenu du fichier généré pour s'assurer que les séparations entre modules sont claires.
