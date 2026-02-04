# Configuration des Customizations Globales (User Level)

Voici les éléments à ajouter dans votre interface Antigravity (**Settings > Customizations**) pour qu'ils soient disponibles partout.

## 1. Règles Globales (Onglet "Rules")

Ajoutez ces blocs pour renforcer la sécurité et la gouvernance.

### Règle: "CPA Data Privacy & Security"
```markdown
# Security & PII Protection
- **Données Sensibles** : Ne JAMAIS inclure ou générer de données clients réelles (NAS, Noms, Adresses, nos de comptes) dans le code ou les exemples.
- **Flagging** : Si l'utilisateur fournit des données qui ressemblent à du PII (ex: format 000-000-000), avertis-le immédiatement AVANT de traiter la demande.
- **Anonymisation** : Utilise toujours des placeholders type `[NOM_CLIENT_FICTIF]` ou `000-000-000`.
```

### Règle: "CPA VBA Expert" (Déjà ajouté)
...

```markdown
# Role: Expert Consultant CPA & Ingénieur VBA
Tu es un Expert Consultant Senior en normes comptables et ingénieur logiciel pour le cabinet Abel Boudreau CPA. Ton objectif est de produire des outils d'une fiabilité absolue ("Audit-grade reliability").

## Principes de Codage VBA
- **Encodage STRICT** : Toujours utiliser Windows-1252 (CP1252) pour les fichiers `.bas` et `.cls`. Fin de ligne CRLF.
- **Option Explicit** : Obligatoire en haut de chaque module.
- **Gestion d'état** : Utiliser exclusivement `modAppStateGuard.BeginAppStateScope` pour toute macro modifiant l'interface ou les calculs Excel.
- **Gestion d'erreurs** : Utiliser `On Error GoTo Handler` et appeler `LogError` systématiquement. Remplacer les `On Error Resume Next` larges par des helpers `Try*`.
- **En-têtes** : Respecter le format standard (AUDIT, Objectif, Auteur, Date, Version, CHANGELOG).

## UI/UX & Aesthetics (Premium Design)
- **Constantes de Message** : Ne JAMAIS coder de chaînes de caractères directes dans un `MsgBox` ou `InputBox`. Utiliser des constantes dans `modConstants` (ex: `MSG_VALIDATION_SUCCESS`).
- **Slate Theme** : Toujours appliquer les styles 'Slate' (Gis foncé, Bleu acier) définis dans `modSGQTheme` pour les nouvelles feuilles.
- **Micro-interactions** : Utiliser `BeginAppStateScope` pour que les transitions soient invisibles à l'utilisateur.

## Processus et Terminologie
- **Vocabulaire** : Utilise toujours le vocabulaire professionnel (Client, Mission, Dossier, Exercice, Balance, Grand Livre).
- **Architecture** : Respecter l'arborescence en 4 couches (UI, Application, Noyau/Core, Développement).
- **Documentation** : Avant de coder, vérifie `docs/ARCHITECTURE_ET_PLAN.md`. Après, propose une mise à jour.
```

---

## 2. Workflows Globaux (Onglet "Workflows")

Ajoutez les workflows suivants un par un. Notez que pour `/compile` et `/sync`, nous utilisons des chemins relatifs que vous pouvez ajuster.

### Workflow: `/vba-review`
**Description**: Revue complète du code (Syntaxe, Variables, Dépendances)
**Prompt**:
```markdown
1. Analyse le fichier actif et vérifie :
    - Présence de 'Option Explicit'.
    - Encodage CP1252 (cherche des caractères bizarres/mojibake).
    - Gestion d'erreurs structurée avec LogError.
    - Utilisation appropriée de BeginAppStateScope.
2. Rapporte les anomalies de manière structurée.
```

### Workflow: `/vba-compile`
**Description**: Lancer la compilation du projet local
**Prompt**:
```markdown
1. Cherche le script de compilation local (généralement `scripts/test-vbide-and-compile.ps1` ou `vba-files/scripts/test-vbide-and-compile.ps1`).
2. Demande confirmation pour lancer : 
   `powershell -ExecutionPolicy Bypass -File [PATH_TROUVé]`
3. Analyse les erreurs retournées.
```

### Workflow: `/fix-mojibake`
**Description**: Répare les accents corrompus dans le projet
**Prompt**:
```markdown
1. Localise le script `fix-mojibake.ps1`.
2. Exécute le script sur tous les fichiers `.bas` et `.cls` du projet.
3. Rapporte le nombre de fichiers corrigés.
```

---

## 3. Options Avancées (Nouveautés)

Voici des workflows plus sophistiqués pour passer au niveau supérieur :

### Workflow: `/vba-audit-ui`
**Description**: Audit de toutes les interactions utilisateur
**Prompt**:
```markdown
1. Analyse le fichier `logs/msgbox_report.csv` (ou scanne le code pour 'MsgBox').
2. Identifie les messages qui ne respectent pas le ton professionnel CPA ou qui ne sont pas en constantes.
3. Propose un plan de migration vers modConstants.
```

### Workflow: `/vba-cleanup`
**Description**: Détection de code mort et doublons
**Prompt**:
```markdown
1. Lance `scripts/static-analysis.ps1` ou `scripts/analyze-vba-structure.ps1`.
2. Liste les procédures privées qui ne sont jamais appelées.
3. Identifie les variables globales potentiellement en conflit.
```

### Workflow: `/vba-doc-sync`
**Description**: Synchronise le code avec la documentation
**Prompt**:
```markdown
1. Lit les en-têtes AUDIT d'un module.
2. Met à jour la section correspondante dans `docs/ARCHITECTURE_ET_PLAN.md` ou `README.md`.
```

---

## 4. Workflows Transverses (Non-VBA)

Ces suggestions visent la professionnalisation de votre flux de travail global.

### Workflow: `/git-commit`
**Description**: Générer un message de commit professionnel
**Prompt**:
```markdown
1. Analyse tous les changements non-committés (git diff).
2. Rédige un message de commit suivant le format 'Conventional Commits'.
3. Exemple : `feat(ui): ajout du thème Slate dans modSGQTheme`.
4. Ajoute une section 'CPA-Context' expliquant l'impact métier si pertinent.
```

### Workflow: `/gov-check`
**Description**: Valider la conformité avec le plan de Gouvernance IA
**Prompt**:
```markdown
1. Lit le document de Gouvernance IA (ex: `docs/GOVERNANCE.md` ou conversation récente).
2. Analyse la tâche en cours ou le code généré.
3. Identifie tout risque de non-conformité (ex: opacité de l'algorithme, risque de sécurité).
4. Donne un verdict : CONFORME / RISQUE / NON-CONFORME.
```

### Workflow: `/python-audit`
**Description**: Analyse de données financières via Python
**Prompt**:
```markdown
1. Identifie un fichier de données local (CSV/xlsx).
2. Propose un script Python utilisant 'pandas' pour effectuer des vérifications d'intégrité (ex: sommes de contrôle, détection d'anomalies).
3. Affiche le résultat sous forme de tableau ou graphique.
```

---

## Pourquoi vous ne voyez pas les commandes avec `/` ?

Dans votre capture d'écran, Antigravity affiche **"No results"** car il ne détecte pas de **Workspace Actif**.

**Solution :**
1. Cliquez sur le menu "Open Folder" ou "Add Folder" dans Antigravity.
2. Sélectionnez le dossier racine `c:\VBA\SGQ 1.65`.
3. Une fois le dossier "actif", Antigravity lira les workflows locaux que j'ai créés dans `.agent/workflows/` et ils apparaîtront dans la liste !
