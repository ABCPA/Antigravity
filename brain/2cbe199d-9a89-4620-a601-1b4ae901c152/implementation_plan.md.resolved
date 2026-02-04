# Résolution des erreurs d'encodage (Mojibake)

Ce plan vise à corriger les problèmes de caractères spéciaux corrompus dans les fichiers VBA du projet SGQ 1.65, en s'assurant que tous les fichiers sont encodés en Windows-1252 (CP1252).

## Proposed Changes

### VBA Modules

#### [MODIFY] [modSGQCreation.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQCreation.bas)
- Conversion vers Windows-1252.
- Nettoyage des chaînes de caractères corrompues (ex: `crǸation` -> `création`).

#### [MODIFY] [modSGQUpdateManager.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQUpdateManager.bas)
- Conversion vers Windows-1252.
- Nettoyage des chaînes de caractères corrompues (ex: `terminǸ` -> `terminé`).

#### [MODIFY] Autres fichiers identifiés par le scan
- Application du même traitement aux fichiers détectés.

## Verification Plan

### Automated Tests
- Exécuter `c:\VBA\SGQ 1.65\scripts\test-vbide-and-compile.ps1` pour valider la compilation après correction.
- Vérifier l'encodage via PowerShell : `Get-Content <file> -Encoding Byte`.

### Manual Verification
- Ouvrir les fichiers dans VS Code (avec l'encodage Windows-1252 configuré) pour valider visuellement les caractères accentués.
