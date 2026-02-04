�@echo off
REM Script pour importer les modules VBA modifiés dans index.xlsm
REM Utilise le script import-into-copy.ps1 du projet

echo ============================================================
echo IMPORTATION DES MODULES VBA DANS INDEX.XLSM
echo ============================================================
echo.

cd /d "C:\VBA\SGQ 1.65.worktrees\worktree-2025-12-16T20-14-20"

echo [1/5] Importation de modSGQInterface.bas...
pwsh -NoProfile -File .\scripts\import-into-copy.ps1 -ModulePath "%cd%\vba-files\Module\modSGQInterface.bas" -ExcelFile "%cd%\index.xlsm"
echo.

echo [2/5] Importation de modWorkbookHandlers.bas...
pwsh -NoProfile -File .\scripts\import-into-copy.ps1 -ModulePath "%cd%\vba-files\Module\modWorkbookHandlers.bas" -ExcelFile "%cd%\index.xlsm"
echo.

echo [3/5] Importation de modSGQAdministration.bas...
pwsh -NoProfile -File .\scripts\import-into-copy.ps1 -ModulePath "%cd%\vba-files\Module\modSGQAdministration.bas" -ExcelFile "%cd%\index.xlsm"
echo.

echo [4/5] Importation de modSGQUtilitaires.bas...
pwsh -NoProfile -File .\scripts\import-into-copy.ps1 -ModulePath "%cd%\vba-files\Module\modSGQUtilitaires.bas" -ExcelFile "%cd%\index.xlsm"
echo.

echo [5/5] Importation de modSGQCreation.bas...
pwsh -NoProfile -File .\scripts\import-into-copy.ps1 -ModulePath "%cd%\vba-files\Module\modSGQCreation.bas" -ExcelFile "%cd%\index.xlsm"
echo.

echo ============================================================
echo IMPORTATION TERMINEE
echo ============================================================
echo.
echo Les 5 modules ont été importés dans index.xlsm
echo.
echo PROCHAINES ETAPES:
echo   1. Ouvrir index.xlsm dans Excel
echo   2. Tester le mode admin (activation/desactivation)
echo   3. Tester ALT+F8 (seules les macros btn* doivent apparaitre)
echo.

pause
�*cascade08"(9e377b08b88e30057baf56c074b3d12f1d22372325file:///c:/VBA/SGQ%201.65/import-modified-modules.bat:file:///c:/VBA/SGQ%201.65