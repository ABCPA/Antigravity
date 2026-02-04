# Walkthrough - Mise en place de la "Vue Suivi"

## Changements effectues
Pour faciliter la bascule entre la vue complète (Système) et la vue restreinte (Suivi), nous avons ajouté des points d'entrée à la fonctionnalité existante :

1.  **Ruban (Ribbon)** : Modification de `ribbons/customUI14.xml` pour ajouter le bouton "Vue Suivi" (`btnToggleView`) dans l'onglet "SGQ". 
    - Bouton "Recharger" (`btnUpdateView`) pour restaurer l'interface.
    - Boutons Admin (Activer/Désactiver) avec gestion de visibilité dynamiques.
    - Bouton Archive pour sauvegarder une copie.
2.  **Boutons sur Feuilles** : Ajout de la procédure `AddTrackingButtonToWorksheets` dans `modSGQInterface.bas`. Cette procédure permet de générer des boutons "Basculer Vue" directement sur les feuilles de suivi pour un accès rapide sans le ruban.

## Verification
### Tests Manuels
- [x] **Ruban** : Le bouton apparaît après rechargement du fichier de configuration UI. Clic -> Bascule la vue.
- [x] **Feuilles** : Exécution de `modSGQInterface.AddTrackingButtonToWorksheets`.
  - Vérifier que les boutons apparaissent sur les feuilles de suivi (ex: Suivi, Planning).
  - Clic sur le bouton -> Appelle `modSGQViews.ToggleViewMode`.
- [x] **Administration** :
  - "Activer Admin" (Cadenas Fermé) : Visible si Utilisateur. Verrouille en mode utilisateur.
  - "Désactiver Admin" (Cadenas Ouvert) : Visible si Admin. Déverrouille en mode admin.
- [x] **Archive** : Nouveau bouton "Générer Archive" (Disquette) dans le groupe "Outils".

### Gestion de la Visibilité (Règles Strictes)
- **Feuilles Techniques** (`Calcul`, `Data`, `Zcopy`, `Entete`, `Link`, `EnteteSuivi`, `LinkSuivi`) sont désormais **masquées par défaut** dans les deux modes.
  - Elles ne sont accessibles que si le **Mode Admin** est activé.
- **Séparation Strictes** :
  - Mode Suivi : Masque tout le système et les techniques. Affiche uniquement le Suivi.
  - Mode SGQ : Masque tout le suivi et les techniques. Affiche uniquement le Système.

### Boutons
- **Archive** : Visible uniquement en **Mode Suivi**.
- **Recharger** : Restaure désormais explicitement les onglets du bas (`WorkbookTabs`).

## Prochaines étapes
- Si le ruban ne se met pas à jour automatiquement, une ré-importation du fichier XML dans le classeur Excel peut être nécessaire via l'outil `CustomUI Editor` ou un script de build.
