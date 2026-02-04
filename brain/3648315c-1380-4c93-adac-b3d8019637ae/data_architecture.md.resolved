# Architecture : Connexion des Données (Data Fabric)

## Objectif
Permettre aux agents (Antigravity, Copilot, NotebookLM) d'accéder aux données réelles de l'expert sans les dupliquer, en respectant la sécurité et la structure existante.

## 1. Cartographie des Sources (Mapping)

| Source Physique | Rôle Expert | Point de Montage Agent (Workspace) | Stratégie d'Accès |
| :--- | :--- | :--- | :--- |
| `.../Doc utiles/Outils/Modèles` | **Base de Connaissances** | `2_Normes_References/Sources_SharePoint` | **RAG Local + NotebookLM** (Upload sélectif) |
| `.../Doc utiles/Références` | **Normes (NIV/NCECF)** | `2_Normes_References/Normes_Officielles` | **NotebookLM** (Cœur normatif) |
| `Z:\CPA` | **Dossiers Clients** | `1_Clients` | **RAG Local** (Lecture seule, sécurisé) |
| `.../Administration/caseware` | **Outils & Templates** | `3_Modeles_Livrables/Caseware` | **Copie/Template Local** |

## 2. Stratégie NotebookLM (Cloud Sécurisé)
Utilisé UNIQUEMENT pour les données non confidentielles (Normes, Guides, Procédures internes).

- **Notebook "CPA Reference Master"** :
    - Contenu : Manuels de l'Ordre, Lois fiscales, Modèles de lettre de mission vierges.
    - Action : Uploader les PDF depuis `Doc utiles/Références`.

## 3. Stratégie Locale (RAG & Scripts)
Utilisée pour les données confidentielles (Clients).

- **Configuration `scripts/ingestion-config.json`** :
    - Définir `Z:\` comme "Remote Source".
    - Le script d'ingestion ne *copiera* pas tout, mais créera un index sur les métadonnées.
    - L'accès complet (lecture PDF) se fait à la demande de l'agent (Lazy Loading).

## 4. Plan d'Action Immédiat

1.  **Créer les Raccourcis (Symlinks)** dans le Workspace pour éviter la duplication.
2.  **Configurer le "Notebook CPA"** : Vous devrez uploader manuellement les PDF clés (je vous fournirai la liste).
3.  **Lancer l'Indexation Légère** : Scanner les noms de dossiers clients sur `Z:\` pour que l'agent connaisse la liste des clients.

## Questions pour validation
- Confirmez-vous que `Z:\CPA` est la racine des clients ? (Vue l'exploration : OUI).
- Acceptez-vous de créer des raccourcis Windows (`.lnk`) dans le workspace vers ces dossiers ?
