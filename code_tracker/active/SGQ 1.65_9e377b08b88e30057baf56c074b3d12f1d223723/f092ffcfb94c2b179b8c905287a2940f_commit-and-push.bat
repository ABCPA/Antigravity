ç@echo off
REM Script pour commiter et pousser les modifications sur main

echo ============================================================
echo COMMIT ET PUSH DES MODIFICATIONS SUR MAIN
echo ============================================================
echo.

cd /d "C:\VBA\SGQ 1.65.worktrees\worktree-2025-12-16T20-14-20"

echo [1/4] Verification du statut Git...
git status --short
echo.

echo [2/4] Ajout des fichiers modifies...
git add vba-files\Module\modSGQInterface.bas
git add vba-files\Module\modWorkbookHandlers.bas
git add vba-files\Module\modSGQAdministration.bas
git add vba-files\Module\modSGQUtilitaires.bas
git add vba-files\Module\modSGQCreation.bas
git add import-modified-modules.bat
git add PREVENT_ENCODING_ISSUES.md
git add MANUAL_CHANGES_GUIDE.md
git add ENCODING_FIX_README.md
git add ENCODING_CORRUPTION_FIX.txt
git add MODIFICATIONS_SUMMARY.txt
git add restore-from-git.bat
git add sanitize-modified-files.bat
echo   OK - Fichiers ajoutes
echo.

echo [3/4] Creation du commit...
git commit -m "feat: Mode plein ecran et masquage macros ALT+F8" -m "MODIFICATIONS PRINCIPALES:" -m "- Replace SHOW.TOOLBAR par DisplayFullScreen pour le mode admin" -m "- Ajoute attributs VB_ProcData.VB_Invoke_Func pour cacher les macros non-btn dans ALT+F8" -m "- Corrige encodage CP1252 (Windows-1252) pour tous les fichiers modifies" -m "" -m "FICHIERS MODIFIES:" -m "- modSGQInterface.bas: TryUpdateInterfaceView() utilise DisplayFullScreen" -m "- modWorkbookHandlers.bas: HandleWorkbookBeforeClose() utilise DisplayFullScreen" -m "- modSGQAdministration.bas: Attributs VBA ajoutes (10 procedures)" -m "- modSGQUtilitaires.bas: Attributs VBA ajoutes (6 procedures)" -m "- modSGQCreation.bas: Attributs VBA ajoutes (10 procedures)" -m "" -m "TOTAL: 59 procedures cachees de ALT+F8, seules les macros btn* restent visibles" -m "" -m "SCRIPTS AJOUTES:" -m "- import-modified-modules.bat: Import automatique des modules" -m "- restore-from-git.bat: Restauration en cas de corruption" -m "- sanitize-modified-files.bat: Sanitization avec encodage CP1252" -m "" -m "DOCUMENTATION:" -m "- PREVENT_ENCODING_ISSUES.md: Guide pour eviter les problemes d'encodage" -m "- MANUAL_CHANGES_GUIDE.md: Guide complet des modifications" -m "- ENCODING_FIX_README.md: Procedure de correction d'encodage" -m "" -m "Date: 2025-12-16"

if %errorlevel% equ 0 (
    echo   OK - Commit cree
) else (
    echo   ERREUR - Echec du commit
    pause
    exit /b 1
)
echo.

echo [4/4] Push vers origin/main...
git push origin main

if %errorlevel% equ 0 (
    echo   OK - Push reussi!
    echo.
    echo ============================================================
    echo SUCCES - Modifications poussees sur main
    echo ============================================================
) else (
    echo   ERREUR - Echec du push
    echo.
    echo ============================================================
    echo ERREUR - Verifiez les messages ci-dessus
    echo ============================================================
)

echo.
pause
ç*cascade08"(9e377b08b88e30057baf56c074b3d12f1d2237232-file:///c:/VBA/SGQ%201.65/commit-and-push.bat:file:///c:/VBA/SGQ%201.65