# Tâches - Phase 30 : Validation Finale et Intégration

- [ ] **Validation Technique**
    - [ ] Exécuter le script de compilation (`test-vbide-and-compile.ps1`) pour vérifier l'intégrité du code VBA.
    - [ ] Vérifier l'absence de références circulaires (`modDiagnostics.ScanCircularDependencies`).
- [ ] **Synchronisation**
    - [ ] Importer tous les modules dans le classeur Excel maître (`Start-VbaAction.ps1 -ImportAll`).
    - [ ] Vérifier que les callbacks du Ruban sont fonctionnels.
- [ ] **Validation Fonctionnelle (Smoke Test)**
    - [ ] Vérifier l'ouverture du formulaire "Nouveau Document".
    - [ ] Vérifier le basculement de vue (Système <-> Suivi).
- [ ] **Clôture**
    - [ ] Mettre à jour `ARCHITECTURE_ET_PLAN.md` avec la Phase 30 complétée.
    - [ ] Créer un tag git ou une release note finale.
