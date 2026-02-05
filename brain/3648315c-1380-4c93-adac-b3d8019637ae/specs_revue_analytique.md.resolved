# Spécification : Moteur de Revue Analytique Automatisée

**Module** : CPA-Analytic-Engine v1.0
**Responsable** : Antigravity (Architecte)
**Statut** : Draft

## 1. Vision Métier
Assisté par ce module, l'expert ne passe plus de temps à *calculer* les variations, mais uniquement à les *analyser* et à les documenter.
Le système digère la Balance de Vérification (Trial Balance) et produit instantanément un tableau de bord des risques et des variations significatives.

## 2. Fonctionnalités Principales

### 2.1 Importation Intelligente (ETL)
- **Source** : Fichiers Excel (Export Caseware, DT Max, Simple Comptable) ou CSV standard.
- **Mapping** : Reconnaissance automatique des comptes (ex: "1xxx" = Actif) via un fichier de mapping configurable (`config/account_mapping.json`).
- **Validation** : Vérification que Débit = Crédit à l'import.

### 2.2 Moteur de Calcul (Core)
- **Comparatif** : N vs N-1 (Montant et Pourcentage).
- **Seuil de Matérialité (S.M.)** :
    - Définition auto ou manuelle du S.M. (ex: 1% du Chiffre d'Affaires).
    - Flag automatique pour toute variation > S.M.
- **Ratios Clés** :
    - Liquidité Générale / Immédiate.
    - Rotation des comptes clients (DSO).
    - Marge Brute (vérification de la cohérence des ventes/coûts).

### 2.3 Interface & Rendu (Output)
- **Feuille "LeadSheet_Analytique"** :
    - Colonnes : Compte, Description, Solde N, Solde N-1, Var $, Var %, [!] Alerte.
    - **Zone de Commentaire** : Une colonne vide protégée où l'expert DOIT justifier les écarts flaggés.
- **Graphiques** : Évolution 5 ans (si historique dispo) pour le C.A. et le Net.

## 3. Architecture Technique (VBA)

### 3.1 Nouveaux Modules
- `modAnalyticETL.bas` : Logique d'import et de nettoyage des données.
- `modAnalyticCalc.bas` : Algorithmes de calcul de variations et ratios.
- `modAnalyticReporting.bas` : Génération des feuilles Excel formatées.

### 3.2 Structure de Données (Types)
```vba
Public Type FinancialAccount
    AccountNumber As String
    Description As String
    BalanceCY As Currency ' Current Year
    BalancePY As Currency ' Prior Year
    Group As String ' Actif, Passif, Ventes, etc.
End Type
```

## 4. Plan de Développement (Sprints)

1.  **Semaine 1** : Création du `modAnalyticETL` (Import CSV simple).
2.  **Semaine 2** : Implémentation du calcul N vs N-1 et génération de la feuille de base.
3.  **Semaine 3** : Ajout des Ratios et du Seuil de Matérialité dynamique.

## 5. Critères de Succès (Definition of Done)
- [ ] Import d'une balance de 500 comptes en < 3 secondes.
- [ ] Les variations > S.M. sont surlignées en rouge automatiquement.
- [ ] Le total Actif = Total Passif + Équité est validé avant tout calcul.
