---
description: Nettoyage du système (Brain & Code Tracker)
---

1. Créer le dossier d'archive pour le Brain s'il n'existe pas.
// turbo
2. Déplacer les sessions du Brain de plus de 30 jours vers `brain/archive` (simulation pour l'instant, demande confirmation).
   `Get-ChildItem -Path "brain" -Directory | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Select-Object Name, LastWriteTime`
3. Lister les sessions Code Tracker actives pour revue.
   `Get-ChildItem -Path "code_tracker/active" -Directory | Select-Object Name, LastWriteTime`
4. Proposer le nettoyage des trackeurs orphelins (modification manuelle requise pour l'instant).
