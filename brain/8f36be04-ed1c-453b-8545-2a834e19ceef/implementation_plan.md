# Plan d'Implémentation : Configuration Serveurs MCP

## Objectif
Donner à Antigravity des "super-pouvoirs" via le Model Context Protocol (MCP) pour qu'il puisse interagir directement avec votre système de fichiers, git et le web de manière sécurisée et structurée.

## User Review Required
> [!IMPORTANT]
> L'installation de ces serveurs nécessite `npx` (Node.js). Assurez-vous qu'il est disponible dans votre terminal.

## Proposed Changes

### 1. Création du fichier de configuration racine
Nous allons créer un fichier `mcp.json` standardisé à la racine du projet (`c:\VBA\SGQ 1.65\mcp.json`) qui remplacera ou complétera celui de `.vscode`.

### 2. Serveurs Recommandés
Nous allons configurer les serveurs suivants :

#### [NEW] @modelcontextprotocol/server-filesystem
*   **But** : Permettre à l'agent de lire/écrire dans vos dossiers de projet sans halluciner les chemins.
*   **Config** : Restreint à `c:\VBA\SGQ 1.65` et `d:\VBA` (si applicable).

#### [NEW] @modelcontextprotocol/server-git
*   **But** : Permettre à l'agent de faire des commits, diffs et logs directement (supporte le workflow `/git-commit`).
*   **Avantage** : Plus robuste que les commandes shell pures.

#### [NEW] @modelcontextprotocol/server-fetch
*   **But** : Permettre à l'agent de lire de la documentation en ligne (ex: Microsoft Learn pour VBA) et de convertir le contenu en Markdown optimisé.

#### [MODIFY] m365agentstoolkit
*   **But** : Nettoyer la configuration existante (doublons détectés) et s'assurer qu'elle pointe vers la bonne version.

## Verification Plan

### Automated Tests
*   Lancer la commande `validate-mcp-json.ps1` (à adapter pour le fichier racine).
*   Demander à l'agent : "Liste les fichiers du dossier actuel via MCP".

### Manual Verification
*   Vérifier que les nouveaux outils apparaissent dans la liste des outils disponibles de l'agent.
