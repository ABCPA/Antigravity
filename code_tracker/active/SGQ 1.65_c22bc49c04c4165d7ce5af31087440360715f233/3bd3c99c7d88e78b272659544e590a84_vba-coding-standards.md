�# VBA Coding Standards (SGQ 1.65)

## Encodage et Formatage
- **Encodage STRICT** : Tous les fichiers `.bas`, `.cls`, `.frm` DOIVENT être enregistrés en `Windows-1252` (CP1252) sans BOM.
- **Fins de ligne** : Utiliser exclusivement `CRLF`.
- **Indentation** : 4 espaces (pas de tabulations).

## Structure du Code
- **Option Explicit** : Requis en haut de CHAQUE module.
- **En-têtes** : Chaque procédure publique doit avoir un en-tête documenté selon `docs/VBA_FILE_HEADER_CONVENTIONS.md`.
- **Gestion d'erreurs** : Utiliser `On Error GoTo Handler` et appeler `LogError` dans tous les modules applicatifs.
- **Gestion d'état** : Utiliser le pattern `BeginAppStateScope` (via `modAppStateGuard`) pour toute opération modifiant l'interface ou les calculs Excel.

## Maintenance
- **Manifeste** : Mettre à jour `vba-files/manifest.json` lors de l'ajout ou de la suppression de modules.
- **Refactoring** : Respecter l'architecture en 4 couches définie dans `docs/ARCHITECTURE_ET_PLAN.md`.
�*cascade08"(c22bc49c04c4165d7ce5af31087440360715f2332>file:///c:/VBA/SGQ%201.65/.agent/rules/vba-coding-standards.md:file:///c:/VBA/SGQ%201.65