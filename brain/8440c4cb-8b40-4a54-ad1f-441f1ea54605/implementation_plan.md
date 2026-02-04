# Plan d'Implémentation - Phase 17 : Excellence & Tests

## Objectif
Clôturer les derniers TODOs de refactorisation et mettre en place une stratégie de tests unitaires pérenne.

## Proposed Changes

### 1. Refactorisation Finale (`modSGQCreation.bas`)
- **Problème** : Un `On Error Resume Next` identifié comme "TODO - remplacer par Try* helper" subsiste ligne 354.
- **Action** : Créer un helper `TryGetWorksheet` dans `modExcelUtils` (ou utiliser l'existant) et l'appliquer.

### 2. Infrastructure de Tests Unitaires
- **Existant** : Dossier `xvba_unit_test` quasi-vide (`Test.bas`).
- **Nouveau** :
    - Créer `modUnitTestEngine` : Un mini-framework de test en pur VBA (Asserts, TestRunner).
    - Créer `Test_Utilitaires.bas` : Premiers tests pour `modSGQUtilitaires`.
    - Créer tâche VS Code : "Exécuter les Tests".

## Verification Plan

### Automated Tests
- Le refactoring de `modSGQCreation` ne doit pas changer le comportement (vérification manuelle ou script).
- Les nouveaux tests unitaires doivent passer (Vert).
