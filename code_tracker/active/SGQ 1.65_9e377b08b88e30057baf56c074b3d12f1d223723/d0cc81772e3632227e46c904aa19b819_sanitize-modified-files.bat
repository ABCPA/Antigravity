�@echo off
REM Script pour sanitizer les fichiers VBA modifiés et les encoder en CP1252

echo ========================================
echo SANITIZATION DES FICHIERS VBA EN CP1252
echo ========================================
echo.

cd /d "C:\VBA\SGQ 1.65.worktrees\worktree-2025-12-16T20-14-20"

echo [1/5] Sanitizing modSGQInterface.bas...
pwsh -NoProfile -File .\scripts\sanitize-vba-file-clean.ps1 -Paths "vba-files\Module\modSGQInterface.bas"
echo.

echo [2/5] Sanitizing modWorkbookHandlers.bas...
pwsh -NoProfile -File .\scripts\sanitize-vba-file-clean.ps1 -Paths "vba-files\Module\modWorkbookHandlers.bas"
echo.

echo [3/5] Sanitizing modSGQAdministration.bas...
pwsh -NoProfile -File .\scripts\sanitize-vba-file-clean.ps1 -Paths "vba-files\Module\modSGQAdministration.bas"
echo.

echo [4/5] Sanitizing modSGQUtilitaires.bas...
pwsh -NoProfile -File .\scripts\sanitize-vba-file-clean.ps1 -Paths "vba-files\Module\modSGQUtilitaires.bas"
echo.

echo [5/5] Sanitizing modSGQCreation.bas...
pwsh -NoProfile -File .\scripts\sanitize-vba-file-clean.ps1 -Paths "vba-files\Module\modSGQCreation.bas"
echo.

echo ========================================
echo SANITIZATION TERMINEE
echo ========================================
echo.
echo Tous les fichiers sont maintenant encodés en CP1252 (Windows-1252)
echo avec terminaisons de ligne CRLF.
echo.
echo Des backups ont été créés dans vba-files\backups\
echo.

pause
�*cascade08"(9e377b08b88e30057baf56c074b3d12f1d22372325file:///c:/VBA/SGQ%201.65/sanitize-modified-files.bat:file:///c:/VBA/SGQ%201.65