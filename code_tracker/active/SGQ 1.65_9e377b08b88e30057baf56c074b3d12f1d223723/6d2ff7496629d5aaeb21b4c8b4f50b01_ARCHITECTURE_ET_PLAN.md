��<!-- AUDIT : sauvegarde 20251104_150442 -> backups/20251104_150442/docs/ARCHITECTURE_ET_PLAN.md.bak -->
### **Schéma d'Architecture Cible**

L'objectif est de regrouper les modules par responsabilité, en clarifiant le réle de chacun.

**1. Couche Interface & événements (UI & Events Layer)**
_Gére toutes les interactions avec l'utilisateur et les événements du classeur._
L'objectif est de regrouper les modules par responsabilité, en clarifiant le rôle de chacun.

**1. Couche Interface & Événements (UI & Events Layer)**
_Gère toutes les interactions avec l'utilisateur et les événements du classeur._

- **`modRibbonSGQ`**: Callbacks et logique du Ruban Office.
- **`modSGQUIActionDispatcher`**: Centralise les points d'entrée pour les macros spéciales et les actions déclenchées par les boutons de l'interface utilisateur.
- **`modSGQContextuel`**: Gestion des menus contextuels (clic droit).
- **`modSGQInterface`**: Logique d'affichage de haut niveau (`UpdateInterfaceView`, notifications, etc.).
- **`modWorkbookHandlers`**: Point d'entrée unique pour les événements du classeur (`Workbook_Open`, `_Activate`, etc.) qui délégue aux services appropriés.

### Changement (2025-10-01): Suppression du module `modSGQEvenements`

- Le module `modSGQEvenements` a été retiré. Les gestionnaires déévénements sont désormais centralisés dans:
  - `Class/ThisWorkbook.cls` (déclencheurs déévénements) appelant
  - `Module/modWorkbookHandlers.bas` (logique centralisée: `HandleWorkbookOpen`, `HandleWorkbookBeforeSave`, etc.)
- Le fichier néest plus référencé dans `vba-files/manifest.json`. Un stub minimal reste présent uniquement pour éviter toute réinjection accidentelle; il est vide et ignoré.
- Impact: aucun appel externe requis; les appels doivent utiliser les handlers centralisés.

**2. Couche Services Applicatifs (Application Services Layer)**
_Contient la logique métier principale, orchestrant les téches complexes._
- **`modWorkbookHandlers`**: Point d'entrée unique pour les événements du classeur (`Workbook_Open`, `_Activate`, etc.) qui délègue aux services appropriés.

### Changement (2025-10-01): Suppression du module `modSGQEvenements`

- Impact: aucun appel externe requis; les appels doivent utiliser les handlers centralisés.

**2. Couche Services Applicatifs (Application Services Layer)**
_Contient la logique métier principale, orchestrant les tâches complexes._

- **`modSGQCreation`**: Toute la logique de création de fichiers et de dossiers.
- **`modSGQExport`**: Procédures pour l'exportation de données (PDF, etc.).
- **`modSGQProtection`**: Gestion du verrouillage/déverrouillage des feuilles et du classeur.
- **`modSGQAdministration`**: Logique spécifique au mode Administrateur.
- **`modSGQUpdateManager`**: Mécanismes de mise é jour des modules.
- **`modSGQUpdateManager`**: Mécanismes de mise à jour des modules.

**3. Couche Noyau & Utilitaires (Core & Utilities Layer)**
_Fondations du projet. Contient le code le plus stable, générique et réutilisable._

- **`modSGQUtilitaires` (Module Central)**:
  - **Responsabilité :** Boéte é outils unique pour TOUT le projet.
  - **Contenu :** Fonctions pures et génériques (manipulation de chaénes de caractéres, etc.).
- **`modExcelUtils`**: Fonctions utilitaires pour manipuler les objets Excel (classeurs, feuilles, plages, etc.).
- **`modAppStateGuard`**:
  - **Responsabilité :** Gestionnaire d'état de l'application (`BeginAppStateScope`). Ne doit pas étre modifié.
- **`modConstants`**: Domicile de toutes les constantes `Public` partagées é travers l'application.
  - **Responsabilité :** Boîte à outils unique pour TOUT le projet.
  - **Contenu :** Fonctions pures et génériques (manipulation de chaînes de caractères, etc.).
- **`modExcelUtils`**: Fonctions utilitaires pour manipuler les objets Excel (classeurs, feuilles, plages, etc.).
- **`modAppStateGuard`**:
  - **Responsabilité :** Gestionnaire d'état de l'application (`BeginAppStateScope`). Ne doit pas être modifié.
- **`modConstants`**: Domicile de toutes les constantes `Public` partagées à travers l'application.

**4. Couche Outils de Développement (Developer Tools Layer)**
_Modules utilisés uniquement pour le développement, l'audit et la maintenance. Ne font pas partie de l'application finale._

- **`modDiagnostics`**: Fonctions de diagnostic et d'audit du projet.
- `modSGQRefactor`, `modVBAInspector`, `modTestWorkbookEvents`, `modLegacyAutoRefactor`.

### **Plan de Travail Formel**

Basé sur ce schéma, voici les étapes de refactorisation que je propose :

**Phase 0 : Fondations et Fiabilisation (TERMINé)**

- **étape 0.1 : Simplification de la gestion de fichiers. (TERMINé)**
  - **Action :** La logique de création de dossiers et de fichiers (anciennement dans `CreateSubfolderFile`) a été refactorisée et centralisée dans `modSGQFiles` pour plus de clarté et de robustesse.
- **étape 0.2 : Mise en place d'un import par manifeste. (TERMINé)**
  - **Action :** L'importation des modules VBA est maintenant pilotée par un fichier `manifest.json`, assurant un ordre de chargement déterministe et fiabilisant le processus de mise é jour via `modSGQUpdateManager`.

**Phase 1 : Centralisation des Utilitaires (Prérequis pour la simplification) (TERMINé)**

- **étape 1.1 : Consolider les fonctions utilitaires dupliquées. (TERMINé)**
  - **Action :** `LogError` et `SanitizeFileNamePart` ont été standardisés et centralisés dans `modSGQUtilitaires`. Tous les appels ont été mis é jour.
- **étape 1.2 : Résoudre les doublons de noms de procédures. (TERMINé)**
  - **Action :** Le doublon de la fonction `IsVbideTrusted` a été résolu. La version recommandée par `gemini.md` a été conservée dans `modSGQUtilitaires` et tous les appels ont été mis é jour. Les autres doublons de fonctions privées (`IsSignatureLine`) ont été laissés en place car ils ne provoquent pas d'erreurs de compilation et sont spécifiques é leurs modules.

**Phase 2 : Refactorisation du Code Applicatif (TERMINé)**

- **étape 2.1 : Unifier les procédures de masquage de lignes. (TERMINé)**
  - **Action :** J'ai créé les fonctions génériques `HideEmptyRowsInRange` et `HideEmptyRowsForNamedRanges` dans `modSGQUtilitaires` et j'ai remplacé les appels aux anciennes fonctions `HideEmptyRows...` par des appels é ces nouvelles fonctions centralisées.
- **étape 2.2 : Factoriser les "wrappers" de boutons. (TERMINé)**
  - **Action :** Introduction d'une procédure `ExecuteActionSafely(procedureName)` qui encapsule la logique de `appScope` et de gestion d'erreur. Les macros de boutons ont été simplifiées pour utiliser ce nouveau mécanisme.
- **étape 2.3 : Unifier la gestion d'état dans les gestionnaires d'événements. (TERMINé)**
  - **Action :** Réfracter les gestionnaires d'événements (`HandleWorkbook...`) dans `modWorkbookHandlers` pour utiliser exclusivement le mécanisme `appScope` (`modAppStateGuard`) pour la gestion de l'état de l'application, en supprimant les appels concurrents é `modSGQSettings`.

**Phase 3 : Nettoyage et Vérification (TERMINé)**

- **étape 3.1 : Supprimer les anciennes procédures devenues obsolétes suite é la refactorisation. (TERMINé)**
  - **Action :** Les procédures dépréciées (`SaveWorkbookSettings`, `RestoreWorkbookSettings`) et les helpers de débogage temporaires ont été supprimés de `modSGQUtilitaires`.
- **étape 3.2 :** Exécuter les outils de diagnostic du projet pour confirmer que les changements n'ont pas introduit de régressions. (TERMINé)
- **étape 3.3 :** Normaliser la gestion d'erreurs dans modSGQFiles (TERMINé) : introduction de helpers structurés (EnsureFolderInternal, TrySaveWorkbookAs) et suppression des On Error Resume Next critiques.
  - Résultats (2025-09-23 11:53) : Refactor_Audit = OK pour tous les modules (ajustement automatique de modSGQRefactor uniquement) et DiagnoseSensitiveInstructions sans nouvelle instruction sensible.

**Phase 4 : Modularisation des Utilitaires et Services Fichiers (TERMINé)**

- **étape 4.1 :** Cartographier les dépendances externes de `modSGQUtilitaires` et `modSGQFiles`. (TERMINé)
- **étape 4.2 :** Standardiser l'utilisation du `appScope` et supprimer les wrappers de compatibilité. (TERMINé)
  - **Action :** Remplacement de tous les appels é `SafeOptimizeForBatch` et `SafeRestoreFromBatch` par le pattern `BeginAppStateScope`. Le module `modSGQAppScopeHelpers` a été supprimé.
- **étape 4.3 :** Modularisation des fonctions de diagnostic. (TERMINé)
  - **Action :** Les fonctions `GenerateDefinedNamesAudit`, `ReportDuplicateProcedureNames`, et `ReportProcedureCollisionsFor` ont été déplacées de `modSGQUtilitaires` vers `modDiagnostics`.

**Phase 5 : Finalisation de la Modularisation (TERMINé)**

- **étape 5.1 :** Audit final de `modSGQUtilitaires` pour identifier toute fonction restante pouvant étre déplacée vers un module plus spécifique. (TERMINé)
- **étape 5.2 :** Suppression du module `modSGQFiles` (maintenant vide) du projet et du manifeste `vba-files/manifest.json`. (TERMINé)
- **étape 5.3 :** Mise é jour de la section "Schéma d'Architecture Cible" de ce document pour refléter la nouvelle structure. (TERMINé)

**Phase 6 : Standardisation de la Gestion d'Erreurs (TERMINé)**

- **étape 6.1 :** Remplacer les `On Error Resume Next` non contrélés par des blocs `On Error GoTo Handler` structurés dans les modules applicatifs critiques (ex: `modSGQCreation`, `modSGQExport`, `modSGQAdministration`). (TERMINé)
- **étape 6.2 :** S'assurer que tous les gestionnaires d'erreurs appellent la procédure centralisée `LogError`. (TERMINé)

**Phase 7 : Amélioration de la Qualité du Code (TERMINé)**

- **étape 7.1 :** Appliquer un formatage de code cohérent é l'ensemble des modules. (TERMINé)
- **étape 7.2 :** Auditer et compléter les en-tétes de documentation pour toutes les procédures publiques. (TERMINé)
  - `modExcelUtils.bas` (TERMINé)
  - `modExportProcedures.bas` (TERMINé)
  - `modLegacyAutoRefactor.bas` (TERMINé)
  - `modRibbonGateway.bas` (TERMINé)
  - `modRibbonSGQ.bas` (TERMINé)
  - `modSGQAdministration.bas` (TERMINé)
  - `modSGQContextuel.bas` (TERMINé)
  - `modSGQCreation` (TERMINé)
  - `modSGQDiagnostics.bas` (TERMINé)
  - `modSGQDiagnosticsTools.bas` (TERMINé - **SUPPRIMé**)
  - `modSGQEvenements.bas` (TERMINé)
  - `modSGQExport.bas` (TERMINé)
  - `modSGQFileSystem.bas` (TERMINé)
  - `modSGQFileTemplates.bas` (TERMINé)
  - `modSGQInterface.bas` (TERMINé)
  - `modSGQMacrosSpeciales.bas` (TERMINé - **CONSOLIDé**)
  - `modSGQProtection.bas` (TERMINé)
  - `modSGQRefactor.bas` (TERMINé)
  - `modSGQSettings.bas` (TERMINé - **SUPPRIMé**)
  - `modSGQTrackingBuilder.bas` (TERMINé)
  - `modSGQUpdateManager.bas` (TERMINé)
  - `modSGQUtilitaires.bas` (TERMINé)
  - `modSGQValidation.bas` (TERMINé)
  - `modSGQVBProjectHelpers.bas` (TERMINé)

**Phase 8 : Documentation et Cléture (TERMINé)**

- **étape 8.1 :** Mettre é jour le fichier `gemini.md` et tout autre document de contribution pour refléter les nouvelles normes et la structure du projet. (TERMINé)
- **étape 8.2 :** Créer un `CHANGELOG.md` pour documenter les changements majeurs effectués lors de cette refactorisation. (TERMINé)

**Phase 9 : Suppression des Modules Vides (TERMINé)**

- **étape 9.1 :** Suppression des modules `.bas` dépréciés et des modules de classe de feuille (`.cls`) vides.
  - **Action :** Les modules `modAppContextState.bas`, `modExcelEnvironment.bas`, `modSGQSettings.bas`, `modSGQDiagnosticsTools.bas` ont été supprimés.
  - **Action :** 46 modules de classe de feuille (`SheetXX.cls`) vides ont été supprimés.
- **étape 9.2 :** Mise é jour du `manifest.json` et de `ARCHITECTURE_ET_PLAN.md`.
  - **Action :** Le `manifest.json` a été mis é jour pour retirer les modules supprimés.
  - **Action :** Ce document a été mis é jour pour refléter la suppression des modules.

**Phase 10 : Regroupement des Modules d'Actions UI (TERMINé)**

- **étape 10.1 :** Consolidation de `modSGQMacrosSpeciales.bas` et `modButtonSGQ.bas` en `modSGQUIActionDispatcher.bas`.
  - **Action :** Le nouveau module `modSGQUIActionDispatcher.bas` a été créé.
  - **Action :** Les appels aux procédures ont été mis é jour dans `modSGQContextuel.bas` et `modRibbonSGQ.bas`.
  - **Action :** Les anciens modules `modSGQMacrosSpeciales.bas` et `modButtonSGQ.bas` ont été supprimés.
  - **Action :** Le `manifest.json` a été mis é jour.
  - **Action :** Ce document a été mis é jour pour refléter la consolidation.

**Phase 11 : Optimisation et Amélioration Continue**

- **étape 11.1 :** Amélioration du parseur JSON (`modSGQUpdateManager.ParseManifestJSON`) (TERMINé).
  - **Action :** Remplacement du parseur RegEx par la bibliothéque VBA-JSON pour une analyse robuste du manifeste.
- **étape 11.2 :** Centralisation des chaénes et noms codés en dur (`Magic Strings`) (TERMINé).
  - **Constat :** De nombreux modules utilisent des noms de feuilles, de plages, des préfixes de fichiers ou des messages `MsgBox` directement dans le code.
  - **Action :** Centraliser ces éléments dans `modConstants` pour améliorer la maintenabilité et la lisibilité.
- **étape 11.3 :** Optimisation de l'exécution dynamique de macros (`Application.Run`) (TERMINé).
  - **Constat :** L'utilisation de `Application.Run nomProc` contourne les vérifications é la compilation et peut entraéner des erreurs d'exécution.
  - **Action :** Remplacer ces appels par des appels directs ou un mécanisme de dispatch plus robuste si possible.
- **Etape 11.4 :** Revision de l'utilisation de `On Error Resume Next` dans les boucles (TERMINÉ).
  - **Constat :** L''utilisation de `On Error Resume Next` é l''intérieur de boucles peut masquer des problémes sous-jacents.
  - **Action :** Remplacer les occurrences dangereuses par des helpers `Try*` avec gestion d''erreur structurée.
  - **Modules traites (2025-11-04) :** `modSGQTrackingBuilder.bas` (TryFinalizeTrackingWorkbook/TryDeleteFile, journalisation livrable) ; `modSGQCreation.bas` (BuildClientDeliverables pour orchestrer 0-SGQ et 6-Suivi) ; `JsonConverter.bas` (helpers TryGetArrayBounds/TryGetArrayBound/TryCoerceVariantToString, suppression des On Error Resume Next).
  - **Historique :** `modSGQValidation.bas`, `modSGQAdministration.bas`, `modSGQInterface.bas`, `modAppStateGuard.bas`, `modExcelUtils.bas`, `modRibbonSGQ.bas`, `modSGQUtilitaires.bas`, `modWorkbookHandlers.bas`, `modSGQVBProjectHelpers.bas`, `modVBAInspector.bas`, `modTestWorkbookEvents.bas`, `JsonConverter.bas`.
  - **Modules restants :** validation finale des utilitaires (`modSGQExport.bas`, `modSGQFileSystem.bas`) et des modules annexes si de nouvelles occurrences apparaissent.
- **étape 11.5 :** Optimisation de la création d'objets (`Scripting.FileSystemObject`).
  - **Constat :** L'objet `Scripting.FileSystemObject` est créé plusieurs fois dans `modSGQVBProjectHelpers.bas`.
  - **Action :** Instancier cet objet une seule fois au niveau du module pour une légére optimisation.
- **étape 11.6 :** Consolidation de `modSGQFileTemplates` dans `modSGQFileSystem` (TERMINé).

  - **Constat :** Les fonctions de `modSGQFileTemplates` sont trés liées aux opérations de `modSGQFileSystem`.
  - **Action :** Fonctions de `modSGQFileTemplates` déplacées vers `modSGQFileSystem`, module supprimé et `manifest.json` mis é jour.

- **étape 11.7 : Correction des dépendances manquantes et standardisation de la journalisation des erreurs fatales (TERMINé).**
  - **Constat :** Un appel é une procédure non définie (`TryCreateRibbonGateway`) provoquait une erreur de compilation. La journalisation des erreurs critiques n'était pas standardisée.
  - **Action :** Ajout des procédures `TryCreateRibbonGateway` et `ReportRibbonGatewayError` dans `modRibbonGateway`. Création d'une nouvelle procédure `Public Sub LogFatalError` dans `modSGQUtilitaires` pour centraliser la gestion des erreurs critiques avec notification é l'utilisateur.

**Phase 12 : Consolidation des modules de diagnostic (TERMINé)**

- **étape 12.1 :** Fusionner `modSGQDiagnostics` et `modExportProcedures` dans `modDiagnostics`. (TERMINé)
**Phase 0 : Fondations et Fiabilisation (TERMINÉ)**

- **Étape 0.1 : Simplification de la gestion de fichiers. (TERMINÉ)**
  - **Action :** La logique de création de dossiers et de fichiers (anciennement dans `CreateSubfolderFile`) a été refactorisée et centralisée dans `modSGQFiles` pour plus de clarté et de robustesse.
- **Étape 0.2 : Mise en place d'un import par manifeste. (TERMINÉ)**
  - **Action :** L'importation des modules VBA est maintenant pilotée par un fichier `manifest.json`, assurant un ordre de chargement déterministe et fiabilisant le processus de mise à jour via `modSGQUpdateManager`.

**Phase 1 : Centralisation des Utilitaires (Prérequis pour la simplification) (TERMINÉ)**

- **Étape 1.1 : Consolider les fonctions utilitaires dupliquées. (TERMINÉ)**
  - **Action :** `LogError` et `SanitizeFileNamePart` ont été standardisés et centralisés dans `modSGQUtilitaires`. Tous les appels ont été mis à jour.
- **Étape 1.2 : Résoudre les doublons de noms de procédures. (TERMINÉ)**
  - **Action :** Le doublon de la fonction `IsVbideTrusted` a été résolu. La version recommandée par `gemini.md` a été conservée dans `modSGQUtilitaires` et tous les appels ont été mis à jour. Les autres doublons de fonctions privées (`IsSignatureLine`) ont été laissés en place car ils ne provoquent pas d'erreurs de compilation et sont spécifiques à leurs modules.

**Phase 2 : Refactorisation du Code Applicatif (TERMINÉ)**

- **Étape 2.1 : Unifier les procédures de masquage de lignes. (TERMINÉ)**
  - **Action :** J'ai créé les fonctions génériques `HideEmptyRowsInRange` et `HideEmptyRowsForNamedRanges` dans `modSGQUtilitaires` et j'ai remplacé les appels aux anciennes fonctions `HideEmptyRows...` par des appels à ces nouvelles fonctions centralisées.
- **Étape 2.2 : Factoriser les "wrappers" de boutons. (TERMINÉ)**
  - **Action :** Introduction d'une procédure `ExecuteActionSafely(procedureName)` qui encapsule la logique de `appScope` et de gestion d'erreur. Les macros de boutons ont été simplifiées pour utiliser ce nouveau mécanisme.
- **Étape 2.3 : Unifier la gestion d'état dans les gestionnaires d'événements. (TERMINÉ)**
  - **Action :** Réfracter les gestionnaires d'événements (`HandleWorkbook...`) dans `modWorkbookHandlers` pour utiliser exclusivement le mécanisme `appScope` (`modAppStateGuard`) pour la gestion de l'état de l'application, en supprimant les appels concurrents à `modSGQSettings`.

**Phase 3 : Nettoyage et Vérification (TERMINÉ)**

- **Étape 3.1 : Supprimer les anciennes procédures devenues obsolètes suite à la refactorisation. (TERMINÉ)**
  - **Action :** Les procédures dépréciées (`SaveWorkbookSettings`, `RestoreWorkbookSettings`) et les helpers de débogage temporaires ont été supprimés de `modSGQUtilitaires`.
- **Étape 3.2 :** Exécuter les outils de diagnostic du projet pour confirmer que les changements n'ont pas introduit de régressions. (TERMINÉ)
- **Étape 3.3 :** Normaliser la gestion d'erreurs dans modSGQFiles (TERMINÉ) : introduction de helpers structurés (EnsureFolderInternal, TrySaveWorkbookAs) et suppression des On Error Resume Next critiques.
  - **Résultats (2025-09-23 11:53) :** Refactor_Audit = OK pour tous les modules (ajustement automatique de modSGQRefactor uniquement) et DiagnoseSensitiveInstructions sans nouvelle instruction sensible.

**Phase 4 : Modularisation des Utilitaires et Services Fichiers (TERMINÉ)**

- **Étape 4.1 :** Cartographier les dépendances externes de `modSGQUtilitaires` et `modSGQFiles`. (TERMINÉ)
- **Étape 4.2 :** Standardiser l'utilisation du `appScope` et supprimer les wrappers de compatibilité. (TERMINÉ)
  - **Action :** Remplacement de tous les appels à `SafeOptimizeForBatch` et `SafeRestoreFromBatch` par le pattern `BeginAppStateScope`. Le module `modSGQAppScopeHelpers` a été supprimé.
- **Étape 4.3 :** Modularisation des fonctions de diagnostic. (TERMINÉ)
  - **Action :** Les fonctions `GenerateDefinedNamesAudit`, `ReportDuplicateProcedureNames`, et `ReportProcedureCollisionsFor` ont été déplacées de `modSGQUtilitaires` vers `modDiagnostics`.

**Phase 5 : Finalisation de la Modularisation (TERMINÉ)**

- **Étape 5.1 :** Audit final de `modSGQUtilitaires` pour identifier toute fonction restante pouvant être déplacée vers un module plus spécifique. (TERMINÉ)
- **Étape 5.2 :** Suppression du module `modSGQFiles` (maintenant vide) du projet et du manifeste `vba-files/manifest.json`. (TERMINÉ)
- **Étape 5.3 :** Mise à jour de la section "Schéma d'Architecture Cible" de ce document pour refléter la nouvelle structure. (TERMINÉ)

**Phase 6 : Standardisation de la Gestion d'Erreurs (TERMINÉ)**

- **Étape 6.1 :** Remplacer les `On Error Resume Next` non contrôlés par des blocs `On Error GoTo Handler` structurés dans les modules applicatifs critiques (ex: `modSGQCreation`, `modSGQExport`, `modSGQAdministration`). (TERMINÉ)
- **Étape 6.2 :** S'assurer que tous les gestionnaires d'erreurs appellent la procédure centralisée `LogError`. (TERMINÉ)

**Phase 7 : Amélioration de la Qualité du Code (TERMINÉ)**

- **Étape 7.1 :** Appliquer un formatage de code cohérent à l'ensemble des modules. (TERMINÉ)
- **Étape 7.2 :** Auditer et compléter les en-têtes de documentation pour toutes les procédures publiques. (TERMINÉ)
  - `modExcelUtils.bas` (TERMINÉ)
  - `modExportProcedures.bas` (TERMINÉ)
  - `modLegacyAutoRefactor.bas` (TERMINÉ)
  - `modRibbonGateway.bas` (TERMINÉ)
  - `modRibbonSGQ.bas` (TERMINÉ)
  - `modSGQAdministration.bas` (TERMINÉ)
  - `modSGQContextuel.bas` (TERMINÉ)
  - `modSGQCreation` (TERMINÉ)
  - `modSGQDiagnostics.bas` (TERMINÉ)
  - `modSGQDiagnosticsTools.bas` (TERMINÉ - **SUPPRIMÉ**)
  - `modSGQEvenements.bas` (TERMINÉ)
  - `modSGQExport.bas` (TERMINÉ)
  - `modSGQFileSystem.bas` (TERMINÉ)
  - `modSGQFileTemplates.bas` (TERMINÉ)
  - `modSGQInterface.bas` (TERMINÉ)
  - `modSGQMacrosSpeciales.bas` (TERMINÉ - **CONSOLIDÉ**)
  - `modSGQProtection.bas` (TERMINÉ)
  - `modSGQRefactor.bas` (TERMINÉ)
  - `modSGQSettings.bas` (TERMINÉ - **SUPPRIMÉ**)
  - `modSGQTrackingBuilder.bas` (TERMINÉ)
  - `modSGQUpdateManager.bas` (TERMINÉ)
  - `modSGQUtilitaires.bas` (TERMINÉ)
  - `modSGQValidation.bas` (TERMINÉ)
  - `modSGQVBProjectHelpers.bas` (TERMINÉ)

**Phase 8 : Documentation et Clôture (TERMINÉ)**

- **Étape 8.1 :** Mettre à jour le fichier `gemini.md` et tout autre document de contribution pour refléter les nouvelles normes et la structure du projet. (TERMINÉ)
- **Étape 8.2 :** Créer un `CHANGELOG.md` pour documenter les changements majeurs effectués lors de cette refactorisation. (TERMINÉ)

**Phase 9 : Suppression des Modules Vides (TERMINÉ)**

- **Étape 9.1 :** Suppression des modules `.bas` dépréciés et des modules de classe de feuille (`.cls`) vides.
  - **Action :** Les modules `modAppContextState.bas`, `modExcelEnvironment.bas`, `modSGQSettings.bas`, `modSGQDiagnosticsTools.bas` ont été supprimés.
  - **Action :** 46 modules de classe de feuille (`SheetXX.cls`) vides ont été supprimés.
- **Étape 9.2 :** Mise à jour du `manifest.json` et de `ARCHITECTURE_ET_PLAN.md`.
  - **Action :** Le `manifest.json` a été mis à jour pour retirer les modules supprimés.
  - **Action :** Ce document a été mis à jour pour refléter la suppression des modules.

**Phase 10 : Regroupement des Modules d'Actions UI (TERMINÉ)**

- **Étape 10.1 :** Consolidation de `modSGQMacrosSpeciales.bas` et `modButtonSGQ.bas` en `modSGQUIActionDispatcher.bas`.
  - **Action :** Le nouveau module `modSGQUIActionDispatcher.bas` a été créé.
  - **Action :** Les appels aux procédures ont été mis à jour dans `modSGQContextuel.bas` et `modRibbonSGQ.bas`.
  - **Action :** Les anciens modules `modSGQMacrosSpeciales.bas` et `modButtonSGQ.bas` ont été supprimés.
  - **Action :** Le `manifest.json` a été mis à jour.
  - **Action :** Ce document a été mis à jour pour refléter la consolidation.

**Phase 11 : Optimisation et Amélioration Continue**

- **Étape 11.1 :** Amélioration du parseur JSON (`modSGQUpdateManager.ParseManifestJSON`) (TERMINÉ).
  - **Action :** Remplacement du parseur RegEx par la bibliothèque VBA-JSON pour une analyse robuste du manifeste.
- **Étape 11.2 :** Centralisation des chaînes et noms codés en dur (`Magic Strings`) (TERMINÉ).
  - **Constat :** De nombreux modules utilisent des noms de feuilles, de plages, des préfixes de fichiers ou des messages `MsgBox` directement dans le code.
  - **Action :** Centraliser ces éléments dans `modConstants` pour améliorer la maintenabilité et la lisibilité.
- **Étape 11.3 :** Optimisation de l'exécution dynamique de macros (`Application.Run`) (TERMINÉ).
  - **Constat :** L'utilisation de `Application.Run nomProc` contourne les vérifications à la compilation et peut entraîner des erreurs d'exécution.
  - **Action :** Remplacer ces appels par des appels directs ou un mécanisme de dispatch plus robuste si possible.
- **Etape 11.4 :** Revision de l'utilisation de `On Error Resume Next` dans les boucles (TERMINÉ).
  - **Constat :** L'utilisation de `On Error Resume Next` à l'intérieur de boucles peut masquer des problèmes sous-jacents.
  - **Action :** Remplacer les occurrences dangereuses par des helpers `Try*` avec gestion d'erreur structurée.
  - **Resultats :**
     - `modSGQValidation.bas` - 10 occurrences -> helpers Try* (TrySafeOptimizeForBatch, TryRestoreFromPrev, TryRestoreExcelSettings)
     - `modSGQTrackingBuilder.bas` - 8 occurrences -> 7 helpers Try* (file ops, workbook, sheets)
     - `modSGQAdministration.bas` - 2 occurrences -> TrySetSheetVisibility helper
     - `modSGQInterface.bas` - 2 occurrences -> TrySetWorksheetScrollArea, TrySetColumnsHidden helpers
     - `JsonConverter.bas` - helpers TryGetArrayBounds/TryGetArrayBound/TryCoerceVariantToString -> suppression des On Error Resume Next
     - Modules restants : `modSGQExport.bas`, `modSGQFileSystem.bas` (verifier nouvelles occurrences au fil des merges)
     - `modAppStateGuard.bas` - 5 occurrences acceptables (3 deja dans Try* helpers, 2 lectures isolees courtes)
     - `modExcelUtils.bas` - 1 occurrence acceptable (lecture worksheet isolee avec restauration immediate)
     - `modSGQExport.bas` - 1 occurrence acceptable (cleanup tmpWb.Close isole)
     - `modRibbonSGQ.bas`, `modSGQUtilitaires.bas`, `modWorkbookHandlers.bas` - deja refactorises (commentaires seulement)
     - `modSGQUtilitaires`, `modSGQVBProjectHelpers`, `modVBAInspector`, `modTestWorkbookEvents`, `JsonConverter.bas` - precedemment completes
  - **Conclusion :** Tous les modules ont été audités. Les occurrences dangereuses (boucles, opérations critiques répétées) ont été éliminées via helpers Try*. Les occurrences restantes sont acceptables (courte portée, non-boucle, cleanup isolé).
- **Étape 11.5 :** Optimisation de la création d'objets (`Scripting.FileSystemObject`).
  - **Constat :** L'objet `Scripting.FileSystemObject` est créé plusieurs fois dans `modSGQVBProjectHelpers.bas`.
  - **Action :** Instancier cet objet une seule fois au niveau du module pour une légère optimisation.
- **Étape 11.6 :** Consolidation de `modSGQFileTemplates` dans `modSGQFileSystem` (TERMINÉ).
  - **Constat :** Les fonctions de `modSGQFileTemplates` sont très liées aux opérations de `modSGQFileSystem`.
  - **Action :** Fonctions de `modSGQFileTemplates` déplacées vers `modSGQFileSystem`, module supprimé et `manifest.json` mis à jour.
- **Étape 11.7 : Correction des dépendances manquantes et standardisation de la journalisation des erreurs fatales (TERMINÉ).**
  - **Constat :** Un appel à une procédure non définie (`TryCreateRibbonGateway`) provoquait une erreur de compilation. La journalisation des erreurs critiques n'était pas standardisée.
  - **Action :** Ajout des procédures `TryCreateRibbonGateway` et `ReportRibbonGatewayError` dans `modRibbonGateway`. Création d'une nouvelle procédure `Public Sub LogFatalError` dans `modSGQUtilitaires` pour centraliser la gestion des erreurs critiques avec notification à l'utilisateur.

**Phase 12 : Consolidation des modules de diagnostic (TERMINÉ)**

- **Étape 12.1 :** Fusionner `modSGQDiagnostics` et `modExportProcedures` dans `modDiagnostics`. (TERMINÉ)
  - **Objectif :** Centraliser tous les outils de diagnostic et de développement en un seul module.
  - **Action :** Le contenu de `modSGQDiagnostics` et `modExportProcedures` a été déplacé dans `modDiagnostics`. Un appel de procédure incorrect a été corrigé dans `modRibbonSGQ` et les anciens modules ont été supprimés. `modExcelUtils` a été conservé en tant que module utilitaire distinct.

**Phase 13 : Correction et maintenance**

- **étape 13.4 : Correction du bug 'subscript out of range' lors de la création de classeurs (TERMINé).**
  - **Constat :** Une erreur 'subscript out of range' se produisait si une feuille de modèle spécifiée dans `CreateStandardWP` n'existait pas.
  - **Action :** Ajout d'une vérification dans `modSGQCreation.bas` pour s'assurer que la feuille existe avant de tenter de la copier, avec un message d'erreur clair pour l'utilisateur.
- **étape 13.1 : Correction des erreurs de labels dans `modDiagnostics.bas` (TERMINé).**
  - **Constat :** Plusieurs procédures dans `modDiagnostics.bas` utilisaient des labels de gestion d'erreur (`ErrHandler:`, `Handler:`) qui ne correspondaient pas é l'instruction `On Error GoTo ...`, provoquant une erreur de compilation "Label not defined".
  - **Action :** Les labels des procédures `ExportProceduresListCSV`, `IsProcedureAvailable`, `ProcedureExists`, `TryGetWindowProperty` et `TryGetProcedureStartLine` ont été renommés pour correspondre é leur instruction `GoTo` respective.
- **étape 13.2 : Ajout de fonctions utilitaires dans `modDiagnostics.bas` (TERMINé).**
- **Étape 13.1 : Correction des erreurs de labels dans `modDiagnostics.bas` (TERMINÉ).**
  - **Constat :** Plusieurs procédures dans `modDiagnostics.bas` utilisaient des labels de gestion d'erreur (`ErrHandler:`, `Handler:`) qui ne correspondaient pas à l'instruction `On Error GoTo ...`, provoquant une erreur de compilation "Label not defined".
  - **Action :** Les labels des procédures `ExportProceduresListCSV`, `IsProcedureAvailable`, `ProcedureExists`, `TryGetWindowProperty` et `TryGetProcedureStartLine` ont été renommés pour correspondre à leur instruction `GoTo` respective.
- **Étape 13.2 : Ajout de fonctions utilitaires dans `modDiagnostics.bas` (TERMINÉ).**
  - **Constat :** Des fonctions de conversion et de formatage sécurisées étaient nécessaires pour des développements futurs dans le module de diagnostic.
  - **Action :** Ajout des fonctions privées `TryConvertToString`, `TryFormatValue` et `OutputStr` pour fournir des helpers de conversion de type et de débogage robustes.
- **étape 13.3 : Mise à jour des plages de masquage de lignes vides.**
  - **Action :** Les plages de masquage de lignes vides pour `HideEmptyRows69` et `HideEmptyRows70` ont été ajustées dans `modSGQInterface.bas` pour refléter les nouvelles exigences. (TERMINÉ)

**Phase 15 : Amélioration des Scripts d'Importation**

- [x] **Tâche :** Améliorer `Start-VbaAction.ps1` pour y inclure un mécanisme d'importation de secours (`AddFromString`) et déprécier les anciens scripts (`import-all-modules.ps1`, `import-into-copy.ps1`). (TERMINÉ)

**Phase 15 : Amélioration des Scripts d'Importation**

- [x] **Tâche :** Améliorer `Start-VbaAction.ps1` pour y inclure un mécanisme d'importation de secours (`AddFromString`) et déprécier les anciens scripts (`import-all-modules.ps1`, `import-into-copy.ps1`). (TERMINÉ)

### **Workflow outils & validations**

1. PromptPrevalidation é cadrer la demande et consigner contraintes.
2. Prompt ciblé (RefactorSafe, PerformanceCheck, ReviewVBA\*, etc.) é établir plan d'action.
3. Scripts PowerShell (ba-files\\import-modules-noninteractive.ps1, scripts\\validate-assistant-config.ps1) é synchroniser modules et lancer validations.
4. TestValidate.prompt.md é réunir preuves (encoding, compile, audits).
5. PreCommit.prompt.md é contréle final + staging ciblé.

_Sauvegardes générées : modules dans logs\\vba_backups, classeurs dans logs\\workbooks._

---

### Backlog / Améliorations futures

1. Debugger au démarrage pour références circulaires (TERMINÉ)

- [x] Objectif: détecter et tracer automatiquement les dépendances circulaires entre modules/classes é l'ouverture du classeur.
- [x] Pistes:
  - Hook sur `ThisWorkbook.Open` -> appel `modDiagnostics.ScanCircularDependencies`.
  - Graphe de dépendances (parsing des appels, Application.Run, etc.).
  - Rapport dans `logs/diagnostics/circular_deps_YYYYMMDD.txt` et sortie immédiate.
- [x] Critéres d'acceptation:
  - Aucune alerte si aucun cycle.
  - Liste explicite des cycles (ex.: A -> B -> C -> A) avec pointeurs utiles.

2. Ouverture et mise é jour des fichiers

- [x] Objectif: automatiser l'ouverture des ressources (manifests/configs) et l'application de mises é jour sécurisées.
- [x] Pistes:
  - Utilitaires dans `modSGQFileSystem` et intégration `modSGQUpdateManager` (backups, validations CP1252/CRLF/schéma JSON).
  - Journalisation détaillée, rollback sur échec.
- [x] Critéres d'acceptation:
  - Backup systématique avant modification.
  - Message utilisateur clair et logs détaillés.
1. PromptPrevalidation → cadrer la demande et consigner contraintes.
2. Prompt ciblé (RefactorSafe, PerformanceCheck, ReviewVBA*, etc.) → établir plan d'action.
3. Scripts PowerShell (vba-files\import-modules-noninteractive.ps1, scripts\validate-assistant-config.ps1) → synchroniser modules et lancer validations.
4. TestValidate.prompt.md → réunir preuves (encoding, compile, audits).
5. PreCommit.prompt.md → contrôle final + staging ciblé.

_Sauvegardes générées : modules dans logs\vba_backups, classeurs dans logs\workbooks._

---

- Critères d'acceptation:
  - Backup systématique avant modification.
  - Message utilisateur clair et logs détaillés.


**Phase 14 : Industrialisation du Workspace (TERMINÉ)**

- **Étape 14.1 : Configuration de l'environnement de développement.**
  - [x] Valider et mettre à jour `.editorconfig` pour forcer l'encodage `windows1252` et les fins de ligne `crlf`.
  - [x] Valider et mettre à jour `.vscode/settings.json` pour définir `files.encoding` à `windows1252` et `files.autoGuessEncoding` à `false`.
- **Étape 14.2 : Création d'un hook Git de pre-commit.**
  - [x] Développer un script de hook qui valide l'encodage des fichiers VBA (`.bas`, `.cls`, `.frm`) avant chaque commit.
  - [x] Le hook devra échouer en cas de détection de caractères non-CP1252 ou de BOM UTF-8.
  - [x] Intégrer l'exécution du script `scripts/check-attributes-local.ps1` dans le hook.
- **Étape 14.3 : Documentation des workflows de collaboration.**
  - [x] Créer une checklist de refactorisation (`docs/guides/REFACTORING_CHECKLIST.md`) basée sur le Step 11.4.
  - [x] Rédiger un guide sur l'utilisation des différentes IA (`docs/guides/GUIDE_COLLABORATION_IA.md`).
- **Étape 14.4 : Mise à jour de la documentation projet.**
  - [x] Mettre à jour `README.md` pour inclure la nouvelle politique d'encodage.
  - [x] Ajouter une entrée dans `CHANGELOG.md` pour documenter la mise en place de ces nouvelles gardes-fous.
- **Étape 14.5 : (Optionnel) Amélioration des templates GitHub.**
  - [ ] Mettre à jour le template de Pull Request (`PULL_REQUEST_TEMPLATE.md`) pour inclure une checklist de validation manuelle des IAs.
- **Étape 14.5 : (Optionnel) Amélioration des templates GitHub.**
  - [x] Mettre à jour le template de Pull Request (`PULL_REQUEST_TEMPLATE.md`) pour inclure une checklist de validation manuelle des IAs.

**Phase 16 : Optimisation de l'Espace de Travail & Antigravity (TERMINÉ)**

- [x] **Tâche :** Création d'un espace de travail unifié CPA_Unified.code-workspace intégrant Python, VBA (C:\VBA) et les artefacts Antigravity. (TERMINÉ)
- [x] **Tâche :** Configuration des tâches VS Code pour les prompts de révision standardisés. (TERMINÉ)
- [x] **Tâche :** Activation de l'approbation automatique des commandes (Always Proceed) par l'utilisateur. (TERMINÉ)







**Phase 17 : Excellence & Tests (TERMINÉ)**

- [x] **Refactoring** : Nettoyage final de modSGQCreation.bas (Zéro TODO).
- [x] **Tests Unitaires** : Moteur modUnitTestEngine et premiers tests modTest_Utilitaires.
- [x] **Intégration** : Tâche VS Code 'Run VBA Unit Tests' configurée.

**Phase 18 : Correction Revue Analytique (TERMINÉ)**

- [x] **Bug Fix** : Gestion des bords (empty/single-cell) dans modAnalyticETL, modAnalyticCalc et modAnalyticReporting.
- [x] **Tests** : Ajout de modTest_AnalyticalBug et intégration au test engine.

**Phase 19 : Audit de Cohérence et Nettoyage (TERMINÉ)**

- [x] **Nettoyage** : Suppression des scripts obsolètes (`import-all-modules.ps1`, `import-into-copy.ps1`).
- [x] **Refactoring** : Simplification de `CreateSubfolderFile` and utilisation du manifeste dans `CopyRequiredModulesToTrackingFile` (`modSGQTrackingBuilder.bas`).
- [x] **Audit de Nommage** : Vérification des conventions et correction critique (corruption de fichier).
- [x] **Performance** : Test de charge final sur les opérations système (OK, Ouverture ~25s).
- [x] **Vérification** : Exécution de la suite complète de tests unitaires (Compilation OK via `sgq-perf.ps1`).


**Phase 20 : Modernisation de l'Interface (UI/UX) (TERMINÉ)**

- [x] **Thème** : Création de `modSGQTheme.bas` pour centraliser les styles (Palette 'Slate', Polices).
- [x] **Intégration** : Application automatique du thème dans `modSGQInterface`.
- [x] **Ruban** : Modernisation de `customUI14.xml` (Icônes, Actions) et implémentation des callbacks.

**Phase 21 : Optimisation des Performances d'Affichage (TERMINÉ)**

- [x] **Logiciel** : Optimisation de `modSGQInterface.UpdateInterfaceView` (Wrapper `ScreenUpdating`).
- [x] **Objets** : Optimisation de `modSGQTheme.ApplyThemeToSheet` (`ApplyThemeToSheet` vise `UsedRange` uniquement).

**Phase 22 : Documentation Utilisateur (TERMINÉ)**

- [x] **CPA** : Création du `MANUAL_UTILISATEUR.md` (Guide simplifié).
- [x] **Dev** : Renommage de `QUICK_START.md` en `DEV_QUICK_START.md`.
- [x] **Portail** : Mise à jour du `README.md` avec des liens clairs par profil.

**Phase 23 : Mécanisme de Mise à Jour Sécurisé (TERMINÉ)**

- [x] **Backup** : Implémentation du snapshot global (`SnapshotModules`) avant toute mise à jour.
- [x] **Update** : Création de `TryUpdateFromFolder` pour orchestrer la mise à jour transactionnelle.
- [x] **Nettoyage** : Refactoring de `UpdateModules` pour utiliser le nouveau moteur sécurisé.

**Phase 24 : Simplification de la Génération de Fichiers (TERMINÉ)**

- [x] **Analyse** : Étudier le processus actuel générant deux fichiers (0-SGQ... et 6-Suivi...) et identifier les sources de bugs.
- [x] **Solution Retenue** : Adoption d'une stratégie "Fichier Unique" avec bascule de vue (Vue Suivi / Vue Système) au lieu de générer deux fichiers distincts.
- [x] **Implémentation** :
  - Création de `modSGQViews.bas` pour gérer les états d'affichage.
  - Ajout d'un toggle bouton "Vue SGQ/SGC" dans le Ruban.
  - Nettoyage du code de génération complexe (`modSGQTrackingBuilder`).
  - Ajout de `AddTrackingButtonToWorksheets` pour placer des boutons d'accès rapide.
- [x] **Stabilisation** : Correction des dépendances (`ExecuteActionSafely`, `ScanCircularDependencies`) et nettoyage des doublons/en-têtes.

**Phase 25 : Amélioration Expérience Utilisateur (UX) (TERMINÉ)**

- [x] **Dashboard** : Création d'une page d'accueil ('Accueil') centralisant la navigation (TDM, Suivi, Admin).
- [x] **Polish** : Application de correctifs visuels globaux (Zoom 100%, Masquage quadrillage) pour un rendu "Application".
- [x] **Navigation** : Mise en place de `modSGQViews.bas` pour gérer les transitions de vue (Système vs Suivi).
- [x] **Refinement** : Modernisation UI (Suppression "Appliquer Thème", Menu "Nouveaux Documents") et correctifs anti-flickering.

**Phase 26 : Optimisation du flux de travail IA (TERMINÉ)**

- [x] **Workflows** : Création de commandes slash standardisés (`/review`, `/compile`, `/plan-update`, `/vba-sync`) pour accélérer les interactions avec l'IA.
- [x] **Intégration** : Configuration des scripts sous-jacents (`test-vbide-and-compile.ps1`, etc.) dans les workflows.


**Phase 27 : Audit et Nettoyage Global (Janvier 2026) (TERMINÉ)**

- [x] **Audit** : Vérification de la cohérence du plan vs codebase (ex: `modSGQTrackingBuilder` manquant, `modSGQUpdateManager`).
- [x] **Nettoyage UI** : Suppression ou reconnexion des boutons morts (ex: `btnCreateSubfolderFile_wrapper` -> `CreateClientCopy`).
- [x] **Standardisation** : Application stricte de `appScope` et `LogError` sur les modules récents.
- [x] **Documentation** : Mise à jour du `manifest.json` si nécessaire.
- **Note** : `modSGQTrackingBuilder`, `modSGQFiles`, `modSGQEvenements` sont confirmés obsolètes/supprimés. `modSGQUpdateManager` a été restauré.

**Phase 28 : Fusion des branches distantes (TERMINÉ)**

- [x] **Fusion** : Intégration du correctif `mcp.json` depuis `worktree-2025-12-16T20-14-20`.
- [x] **Documentation** : Récupération de `CI_IMPORT_CHECKLIST.md` depuis `copilot/nuclear-opossum`.

**Phase 29 : Optimisation du Workspace et Sécurisation (Janvier 2026) (TERMINÉ)**

- [x] **Sécurisation** : Mise à jour de `.gitignore` et `CPA_Unified.code-workspace` pour exclure strictement les fichiers de données (`.xlsx`, `.csv`).
- [x] **Nettoyage** : Purge du cache Git pour les fichiers techniques existants (`vba-files/reports/*.csv`).
- [x] **Performance** : Refactorisation critique de `HideEmptyRowsInRange` dans `modExcelUtils.bas` pour utiliser des tableaux en mémoire (`Value2` Array) au lieu d'itérations sur cellules.
- [x] **Stabilisation** : Le script `test-vbide-and-compile.ps1` a été corrigé (Compatibilité PS 5.1).

�� ����*cascade08���� ��Ľ*cascade08ĽŽ Žǽ*cascade08ǽѽ ѽҽ*cascade08ҽþ þľ*cascade08ľ�� ����*cascade08��ſ ſƿ*cascade08ƿǿ ǿȿ*cascade08ȿ˿ ˿̿*cascade08̿ο οѿ*cascade08ѿҿ ҿӿ*cascade08ӿԿ Կֿ*cascade08ֿܿ ܿ޿*cascade08޿߿ ߿�*cascade08�� ��*cascade08�� ��*cascade08�� ��*cascade08�� ���*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ����*cascade08���� ���� *cascade08����*cascade08���� *cascade08����*cascade08���� *cascade08����*cascade08���� *cascade08����*cascade08���� *cascade08����*cascade08���� *cascade08����*cascade08���� *cascade08����*cascade08���� *cascade08"(9e377b08b88e30057baf56c074b3d12f1d22372326file:///c:/VBA/SGQ%201.65/docs/ARCHITECTURE_ET_PLAN.md:file:///c:/VBA/SGQ%201.65