�	@echo off
REM Script pour restaurer les fichiers corrompus depuis Git et refaire les modifications

echo =========================================================
echo RESTAURATION DES FICHIERS VBA CORROMPUS
echo =========================================================
echo.

cd /d "C:\VBA\SGQ 1.65.worktrees\worktree-2025-12-16T20-14-20"

echo [ETAPE 1] Restauration depuis Git...
echo.

git checkout vba-files\Module\modSGQInterface.bas
git checkout vba-files\Module\modWorkbookHandlers.bas
git checkout vba-files\Module\modSGQAdministration.bas
git checkout vba-files\Module\modSGQUtilitaires.bas
git checkout vba-files\Module\modSGQCreation.bas

echo.
echo ✓ Fichiers restaurés depuis Git
echo.
echo =========================================================
echo ATTENTION - MODIFICATIONS A REFAIRE
echo =========================================================
echo.
echo Les fichiers ont été restaurés à leur état d'origine.
echo.
echo Les modifications doivent être refaites MANUELLEMENT
echo dans l'éditeur VBA avec l'encodage CP1252.
echo.
echo Voir le fichier MANUAL_CHANGES_GUIDE.md pour les détails.
echo.

pause
�	*cascade08"(9e377b08b88e30057baf56c074b3d12f1d2237232.file:///c:/VBA/SGQ%201.65/restore-from-git.bat:file:///c:/VBA/SGQ%201.65