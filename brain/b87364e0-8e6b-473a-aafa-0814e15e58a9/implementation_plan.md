# Plan de Révision Exhaustive SGQ 1.65

Ce plan définit une stratégie systématique pour auditer l'ensemble du projet afin de garantir une fiabilité absolue ("Audit-grade reliability") conforme aux normes CPA du cabinet Abel Boudreau.

## Objectifs de la révision
- Confirmer l'intégrité technique (compilation, synchronisation).
- Valider la conformité aux standards (encodage, en-têtes, gestion d'erreurs).
- Assurer le professionnalisme de l'interface (normes CPA).
- Nettoyer le code mort et les redondances.

## Nouveaux Outils de Documentation et Contrôle

### 1. Registre de Connaissances (VBA_AUDIT_LOG.md) [NEW]
Un fichier de suivi dynamique sera créé pour capitaliser l'intelligence du code :
- **Variables & Dépendances** : Cartographie des déclarations critiques.
- **Appels Modules** : Traçabilité des dépendances inter-modules.
- **Registre d'Erreurs** : Historique des anomalies détectées et résolues.
- **Log de Modifications** : Justification des suppressions et changements structurels.

### 2. Protocole de Boucle de Révision Itérative
Pour chaque correction (ex: doublons, erreurs de syntaxe) :
- **Action** : Correction ciblée.
- **Révision** : Re-scan immédiat via l'agent QA ou le script de compilation.
- **Itération** : Jusqu'à 3 itérations (ou "Zéro Erreur") avant de considérer la tâche complétée.

## Stratégie de Révision (Phasée) [MISE À JOUR]

### Phase 1 : Initialisation et Registre
1. **Création du Registre** : Initialiser `VBA_AUDIT_LOG.md` avec l'état actuel.
2. **Validation du Manifeste** : `/vba-sync` pour synchroniser `manifest.json`.
3. **Audit de Gouvernance** : `/gov-check` (CP1252, backups).

### Phase 2 : Analyse Statique et Nettoyage (Loop)
Action itérative sur les anomalies :
1. **Audit QA** : `/qa-audit` pour l'analyse statique.
2. **Correction & Loop** : Appliquer les corrections, puis ré-auditer immédiatement.
3. **Documentation** : Noter les suppressions et changements dans le log.

### Phase 3 : Standards Professionnels et UI
1. **Audit UI** : `/vba-audit-ui` pour le ton CPA.
2. **Revue d'Architecture** : `/review ALL ALL` (AppStateGuard, 4 couches).

### Phase 4 : Documentation Finale
1. **Synchronisation Doc** : `/vba-doc-sync`.
2. **Fermeture du Registre** : État final documenté dans le log.
3. **Mise à jour du Plan global** : `/plan-update`.

## Plan de Vérification

### Tests Automatisés
- `scripts\test-vbide-and-compile.ps1` : Doit retourner Exit Code 0.
- `scripts\vba-analyze.ps1` : Doit produire un rapport JSON sans erreurs critiques.
- `/test-all` : Lancer la suite complète de tests unitaires (`modUnitTestEngine`).

### Vérification Manuelle
- Ouverture du fichier `index.xlsm` pour valider l'absence d'erreurs au démarrage (`Workbook_Open`).
- Test visuel du Ruban pour confirmer que les callbacks sont fonctionnels.
