# Spécification : Gardien de Contrôle Qualité (CQ)

**Module** : CPA-Compliance-Guardian v1.0
**Objectif** : Garantir qu'aucun dossier ne soit clôturé s'il manque des éléments critiques (Norme de Contrôle Qualité).

## 1. Vision Métier
L'outil agit comme un "Second Associé" virtuel. Il scanne le dossier Excel et vérifie la présence des signatures et documents obligatoires avant de permettre l'exportation des États Financiers.

## 2. Fonctionnalités Principales

### 2.1 Checklist Dynamique (Smart Checklist)
Génère une liste de points de contrôle basée sur les données du dossier :
- Si `Stocks > 0` -> Ajoute "Présence à l'inventaire effectuée ?".
- Si `Dette > 0` -> Ajoute "Confirmations bancaires reçues ?".
- Si `Mission = Audit` -> Ajoute "Lettre d'affirmation de la direction signée ?".

### 2.2 Validateur de Signatures (Sign-off Tracker)
Vérifie les métadonnées des feuilles maîtresses (Lead Sheets) :
- **Préparé par** (Initiales + Date) : Requis pour TOUTES les feuilles.
- **Revu par** (Initiales + Date) : Requis pour les feuilles à risque élevé.

### 2.3 Tableau de Bord CQ (Dashboard)
Une feuille "Contrôle_Qualité" présentant :
- État d'avancement global (%).
- Liste des manquants bloquants (Rouge).
- Liste des avertissements (Jaune).

## 3. Architecture Technique (VBA)

### 3.1 Nouveaux Modules
- `modCQRules.bas` : Moteur de règles (Si X alors Y requis).
- `modCQScanner.bas` : Parcourt les feuilles pour vérifier les cellules de signature (ex: Cellule "H1" = Préparateur).
- `modCQReporting.bas` : Génère le rapport CQ.

### 3.2 Configuration (JSON)
- `cq_rules.json` : Fichier externe définissant les règles pour éviter de coder en dur.
  *(Ex: "AccountGroup:STOCKS" -> "Check:Procédure Inventaire")*

## 4. Plan de Développement
1.  **Semaine 1** : Création du Scanner de Signatures (`modCQScanner`).
2.  **Semaine 2** : Implémentation du Moteur de Règles (`modCQRules`).
3.  **Semaine 3** : Intégration dans le flux de clôture (Empêcher PDF si CQ échec).

## 5. Critères de Succès
- [ ] Détecte infailliblement une feuille non signée.
- [ ] Bloque la procédure "Générer États Financiers" si erreurs critiques.
- [ ] Temps de scan < 2 secondes pour un classeur de 20 feuilles.
