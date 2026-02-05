# Walkthrough : Rules et Workflows pour SGQ 1.65

J'ai implémenté un ensemble complet de règles et de flux de travail pour garantir la stabilité et la cohérence de votre projet VBA.

## Changements Effectués

### 1. Nouvelles Rules (`.agent/rules/`)
Ces règles guident le comportement d'Antigravity pour qu'il respecte vos standards techniques :
- **[vba-coding-standards.md](file:///c:/VBA/SGQ%201.65/.agent/rules/vba-coding-standards.md)** : Impose l'encodage CP1252, l'usage de `LogError`, `BeginAppStateScope` et `Option Explicit`.
- **[process-rules.md](file:///c:/VBA/SGQ%201.65/.agent/rules/process-rules.md)** : Impose la mise à jour systématique du plan d'architecture et la validation par compilation.

### 2. Nouveaux Workflows (`.agent/workflows/`)
Vous pouvez maintenant utiliser les raccourcis suivants (slash commands) :
- **`/fix-mojibake`** : Convertit et répare les accents corrompus.
- **`/test-all`** : Lance la compilation ET les tests unitaires.
- **`/refactor-feature`** : Guide étape par étape pour migrer du code vers l'architecture cible.
- **`/compile`** et **`/vba-sync`** restent disponibles et sont optimisés.

### 3. Workflows Avancés & Transverses
Pour aller plus loin, j'ai aussi pré-créé ces workflows locaux :
- **`/vba-audit-ui`** : Vérifie que vos MsgBox respectent le ton 'CPA'.
- **`/git-commit`** : Rédige vos messages de commit comme un pro.
- **`/gov-check`** : Valide la sécurité et la gouvernance de vos changements.
- **`/python-audit`** : Analyse vos données Excel via Python.

## Impact pour le Projet
- **Moins de bugs d'encodage** : L'IA ne générera plus de fichiers UTF-8 par erreur.
- **Qualité de code accrue** : Standardisation de la gestion d'erreurs et d'état.
- **Autonomie de l'IA** : Elle sait maintenant qu'elle doit mettre à jour `ARCHITECTURE_ET_PLAN.md` toute seule.

---
> [!TIP]
> Vous pouvez maintenant demander à Antigravity : "Lanse le workflow /test-all pour vérifier l'état du projet".
