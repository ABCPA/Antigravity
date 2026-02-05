# Plan de résolution des doublons QA (SGQ 1.65)

Ce plan vise à éliminer les conflits de noms identifiés par l'audit QA du 2026-02-04 pour garantir une fiabilité de niveau audit ("Audit-grade").

## User Review Required

> [!IMPORTANT]
> - Les fonctions de visibilité des feuilles seront centralisées dans `modSGQProtection`.
> - La fonction `FolderExists` de `modSGQAuditLog` sera supprimée au profit de `modSGQFileSystem`.
> - Les fonctions `CreateIssue` (privées) seront renommées pour éviter les collisions sémantiques.

## Proposed Changes

### [Services Applicatifs]

#### [MODIFY] [modSGQAdministration.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQAdministration.bas)
- Supprimer les déclarations redondantes de `TrySetSheetVisibility` et `SetAdminSheetsVisibility`.
- Vérifier que les appels utilisent bien les versions publiques de `modSGQProtection`.

#### [MODIFY] [modSGQAuditLog.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQAuditLog.bas)
- Supprimer la fonction `Private Function FolderExists`.
- Mettre à jour les appels internes pour utiliser `modSGQFileSystem.FolderExists`.

#### [MODIFY] [modCQRules.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modCQRules.bas)
- Renommer `CreateIssue` en `CreateCQRulesIssue`.

#### [MODIFY] [modInspectionAudit.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modInspectionAudit.bas)
- Renommer `CreateIssue` en `CreateInspectionIssue`.

## Verification Plan

### Automated Tests
- Exécuter `/vba-compile` (scripts\test-vbide-and-compile.ps1) pour valider l'intégrité du projet.
- Exécuter un nouveau scan QA (`scripts\vba-analyze.ps1`) pour confirmer la disparition des doublons.

### Manual Verification
- Demander à l'utilisateur de vérifier l'accès au mode Administrateur (qui utilise `SetAdminSheetsVisibility`).
- Vérifier la génération des logs d'audit.
