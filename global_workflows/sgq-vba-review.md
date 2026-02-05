---
description: Revue complète du code (Syntaxe, Variables, Dépendances)
---

1. Analyse le fichier actif et vérifie :
    - Présence de 'Option Explicit'.
    - Encodage CP1252 (cherche des caractères bizarres/mojibake).
    - Gestion d'erreurs structurée avec LogError.
    - Utilisation appropriée de BeginAppStateScope.
2. Rapporte les anomalies de manière structurée.