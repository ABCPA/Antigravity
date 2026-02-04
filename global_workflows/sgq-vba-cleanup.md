---
description: Détection de code mort et doublons
---

// turbo
1. Lance l'analyse statique :
   `powershell -File scripts/static-analysis.ps1`
2. Liste les procédures privées qui ne sont jamais appelées.
3. Identifie les variables globales potentiellement en conflit.