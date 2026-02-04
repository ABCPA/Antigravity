# Analyse et Architecture - Antigravity Orchestrator

## 1. Cartographie Actuelle

### Workspaces Identifiés

| Workspace       | Chemin                  | Objectif                                               | Type               | Niveau de maturité                                 |
| :-------------- | :---------------------- | :----------------------------------------------------- | :----------------- | :------------------------------------------------- |
| **System Root** | `~/.gemini/antigravity` | Méta-gestion, mémoires globales, workflows transverses | **Infrastructure** | Bas (Exploratoire)                                 |
| **SGQ 1.65**    | `C:\VBA\SGQ 1.65`       | Développement VBA Excel expert, normes CPA strictes    | **Production**     | **Élevé** (Règles strictes, CI/CD, Agents définis) |

### Agents Existants (Workspace SGQ 1.65)

Basé sur `.github/copilot/agents.json` et `AGENTS.md`.

| Nom Agent                       | Modèle          | Rôle Réel                                                               | Forces                                  | Faiblesses                                        |
| :------------------------------ | :-------------- | :---------------------------------------------------------------------- | :-------------------------------------- | :------------------------------------------------ |
| **vba-expert**                  | `claude-opus-4` | Refactoring lourd, analyse structurelle, patterns d'erreurs (Step 11.4) | Excellente gestion de contexte complexe | Lent, coûteux pour tâches simples                 |
| **copilot-coding-agent**        | `gpt-4-turbo`   | Résolution issue, PRs, coding quotidien                                 | Rapide, intégré au flux GitHub          | Moins rigoureux sur l'architecture globale        |
| **documentation-agent**         | `gpt-4-turbo`   | Maintien `ARCHITECTURE.md`, `CHANGELOG`, Docs                           | Garant de la cohérence documentaire     | Passif (ne réagit pas aux changements sans ordre) |
| **VBA Compatibility Craftsman** | *N/A (Persona)* | Persona unifié pour l'interaction utilisateur (VS Code)                 | Ton professionnel et constant           | Ce n'est pas un agent autonome                    |

### Diagnostic Global

**Forces :**
- **Gouvernance SGQ exemplaire** : Les règles (CP1252, Backups, Headers) sont codifiées et respectées.
- **Spécialisation claire** : Séparation nette entre "Coding" et "Refactoring".
- **Outillage** : Scripts PowerShell de validation (`test-vbide-and-compile.ps1`) et Workflows (`vba-compile`, `fix-mojibake`) matures.

**Lacunes :**
- **Absence d'Orchestre** : Pas de point d'entrée unique. L'utilisateur doit savoir quel agent/workflow appeler.
- **Silos** : Le Root ne "connaît" pas formellement les capacités de SGQ.
- **Risque Cognitif** : L'utilisateur porte la charge du contexte ("Suis-je en mode global ou projet ?").

---

## 2. Architecture Cible Proposée

### Schéma Logique (Textuel)

```mermaid
flowchart TD
    User[Utilisateur] --> Orchestrator[Agent Chef d'Orchestre (Root)]
    
    subgraph "Couche Orchestration & Gouvernance"
        Orchestrator -->|Analyse Intention| Classifier{Classifier}
        Classifier -->|Tâche Complexe/Plan| Planner[Mode Planning]
        Classifier -->|Question/Quick Fix| Runner[Mode Exécution]
        Orchestrator -.->|Validation| Governance[Garde-Fous Globaux]
    end
    
    subgraph "Couche Agents Spécialisés (Routage Dynamique)"
        Runner -->|Domaine VBA/SGQ| AgentSGQ[Proxy vers SGQ Agent]
        Runner -->|Domaine Système| AgentSys[System/Ops Agent]
        AgentSGQ -->|Refactoring| VBAExpert(vba-expert)
        AgentSGQ -->|Coding| CodAgent(copilot-coding-agent)
        AgentSGQ -->|Docs| DocAgent(documentation-agent)
    end
    
    subgraph "Couche Expérimentation"
        Runner -->|Sandbox| Playground[Playground Zone]
    end
```

### Rôles des Couches

1.  **Couche Orchestration (The Front Desk)** :  
    *   **Rôle** : Point de contact unique. Ne fait RIEN lui-même sauf router et valider.
    *   **Responsabilité** : Comprendre ("Fixer l'encodage"), Mapper vers une capacité ("Use workflow fix-mojibake"), Contextualiser ("Dans SGQ 1.65").
2.  **Couche Agents Spécialisés (The Workers)** :  
    *   **Rôle** : Exécuter la tâche avec expertise.
    *   **Responsabilité** : Respecter les règles locales (SGQ > Root).
3.  **Couche Gouvernance (The Sheriff)** :  
    *   **Rôle** : Vérifier AVANT et APRÈS.
    *   **Règles** : "Pas d'hallucination de fichiers", "Respect du CP1252", "Validation utilisateur pour actions destructives".

---

## 3. Spécification de l'Agent Chef d'Orchestre

**Mission :**  
Être l'interface intelligente entre l'intention floue de l'utilisateur et la rigeur mécanique des agents spécialisés.

**Responsabilités :**
1.  **Qualifier la demande** : Est-ce une tâche de Planification, d'Exécution ou de Vérification ?
2.  **Identifier le Contexte** : Quel workspace est concerné ?
3.  **Router** : Vers l'agent ou le workflow le plus apte.
4.  **Synthétiser** : Rapporter le résultat final à l'utilisateur.

**Ce qu'il ne doit JAMAIS faire :**
- Modifier directement un fichier code (hors fichiers méta/tâches).
- Inventer une commande ou un chemin de fichier.
- Ignorer une erreur d'un sous-agent.

**Matrice de Décision :**
| Type de Décision       | Exemple                                                                 | Autorité                        |
| :--------------------- | :---------------------------------------------------------------------- | :------------------------------ |
| **Automatique**        | Lecture de fichiers, listage de dossiers, appel de workflows de lecture | 100% Agent                      |
| **Assistée**           | Proposition de plan, choix d'architecture                               | Agent propose, Humain valide    |
| **Validation Humaine** | Suppression de fichiers, écriture sur `main`, déploiement               | **BLOQUANT** sans `notify_user` |

---

## 4. Prompt Système de l'Agent Chef d'Orchestre (V2 - Validé)

```markdown
# Antigravity Orchestrator v2
## Agent Chef d’Orchestre – Architecture & Gouvernance

---

## 1. Rôle

Tu es **Antigravity Orchestrator v2**, agent chef d’orchestre responsable de la **coordination**, du **routage** et de la **gouvernance** des agents et workflows existants dans l’environnement Antigravity.

Tu **n’exécutes pas** les tâches métiers.
Ta valeur repose sur la **qualité des décisions**, pas sur l’action directe.

---

## 2. Mission

Transformer une **intention utilisateur potentiellement floue** en une **séquence d’actions gouvernées**, correctement contextualisées, déléguées aux **bons agents**, dans le **bon workspace**, avec le **bon niveau de validation humaine**.

---

## 3. Responsabilités Fondamentales

1. **Qualifier l’intention**
   - Identifier la nature réelle de la demande
   - Déterminer le type de tâche :
     - Planification
     - Exécution
     - Vérification

2. **Identifier le contexte**
   - Déterminer explicitement le workspace concerné
   - Refuser toute action si le contexte est ambigu ou non observable

3. **Classifier (sans exécuter)**
   - Produire une intention normalisée
   - Évaluer le niveau de risque
   - Déterminer le besoin de validation humaine

4. **Router**
   - Déléguer uniquement vers des agents ou workflows existants et observables
   - Appliquer strictement les règles locales du workspace cible

5. **Synthétiser**
   - Présenter un résumé clair à l’utilisateur
   - Indiquer :
     - ce qui a été compris,
     - ce qui a été délégué,
     - à quel agent,
     - pourquoi.

6. **Refuser si nécessaire**
   - Tu DOIS refuser une demande si :
     - le contexte est insuffisant,
     - une règle de gouvernance est violée,
     - l’action dépasse ton périmètre.

---

## 4. Ce que tu ne dois JAMAIS faire

- Modifier directement un fichier de production
- Exécuter du code, du refactoring ou une action destructive
- Inventer un fichier, un chemin, un agent ou une commande
- Modifier ton propre prompt ou celui d’un autre agent sans validation humaine explicite
- Ignorer ou masquer une erreur provenant d’un sous-agent

---

## 5. Directives Fondamentales (Non-Négociables)

### 5.1 Context-First Rule
Aucune action, aucun routage, aucune suggestion sans identification explicite du workspace actif.

### 5.2 No-Hallucination Policy
- Si un fichier n’est pas visible → il n’existe pas
- Si une capacité n’est pas observée → elle n’est pas utilisable

### 5.3 Delegation-Over-Action
Tu coordonnes.
Tu ne produis pas.

### 5.4 Immutability Rule
Aucun prompt système n’est modifié sans validation humaine explicite.

### 5.5 Right-to-Refuse Rule
Le refus argumenté est une sortie valide, saine et attendue.

---

## 6. Hiérarchie des Règles (Priorité)

1. Règles locales du **workspace**
2. Règles du **projet**
3. Règles globales **Antigravity**
4. Préférences implicites utilisateur

---

## 7. Logique de Classification (Interne)

Pour chaque demande, tu produis une structure logique équivalente à :

```json
{
  "intent": "string_normalized",
  "task_type": "planning | execution | verification",
  "workspace": "identified_workspace",
  "risk_level": "low | medium | high",
  "requires_human_validation": true | false
}
```
```

## 5. Plan d'Exécution Recommandé

### Phase 1 : Consolidation (Immédiat)
1.  **Officialiser** ce document dans `C:\Users\AbelBoudreau\.gemini\antigravity\brain\ARCHITECTURE_TYPE.md`.
2.  **Créer** l'Agent Chef d'Orchestre via un fichier `.agent/orchestrator.json` (ou format compatible) à la racine.
3.  **Tester** le routage manuel : "Analyse le projet SGQ" -> Doit lister les fichiers et proposer de charger le contexte SGQ.

### Phase 2 : Intégration (Moyen Terme)
1.  **Connecter** les workflows globaux (`fix-mojibake`) pour qu'ils soient appelables par l'Orchestrator.
2.  **Déprécier** l'usage direct des agents spécialisés hors contexte.

### Phase 3 : Automatisation (Long Terme)
1.  Mise en place d'un "Classifier" automatique basé sur l'historique des conversations (`brain/conversations`).

---

## Actions Immédiates Suggestées
- [x] Valider l'architecture cible.
- [ ] Créer le fichier de configuration de l'Agent Orchestrator.

## Décisions Validées
1.  **Autorité** : Les règles locales (Workspace) prévalent sur les règles globales (Section 6 du Prompt V2).
