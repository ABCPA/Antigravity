�# Process Rules (SGQ 1.65)

## Avant de Commencer
- **Consulter le Plan** : Lire `docs/ARCHITECTURE_ET_PLAN.md` pour comprendre le contexte architectural.
- **Mettre à jour ARCHITECTURE_ET_PLAN.md** : Enregistrer chaque étape majeure complétée dans le journal des modifications de ce fichier.

## Pendant le Développement
- **Fichiers de Travail** : Si possible, travailler sur les fichiers `_worktree.bas` pour les modifications importantes avant de les fusionner dans les modules principaux.
- **Sauvegardes** : Antigravity doit demander une sauvegarde (via script ou copie manuelle) environ tous les 4 modules traités.

## Après les Modifications
- **Vérification de Compilation** : Lancer systématiquement `scripts\test-vbide-and-compile.ps1` (ou via le workflow `/compile`).
- **Tests Unitaires** : Exécuter les tests pertinents via `modUnitTestEngine` si la modification impacte la logique métier.
�*cascade08"(c22bc49c04c4165d7ce5af31087440360715f23327file:///c:/VBA/SGQ%201.65/.agent/rules/process-rules.md:file:///c:/VBA/SGQ%201.65