# Implementation Plan - UX Improvements & Global Import

Standardize the user experience with a modern 'Accueil' dashboard and apply visual polish across all sheets. Finally, relaunch the global import of VBA modules.

## User Review Required

> [!IMPORTANT]
> The 'Accueil' sheet will be created and set as the first sheet. All existing sheets will have their gridlines hidden and zoom set to 100% to create a professional "Application" feel.

## Proposed Changes

### [VBA Infrastructure]

#### [MODIFY] [modWorkbookHandlers](file:///C:/VBA/SGQ%201.65/vba-files/Module/modWorkbookHandlers.bas)
- Hook `SetupDashboard` into `HandleWorkbookOpen` to ensure the landing page is prepared and activated on startup.

#### [MODIFY] [modSGQInterface](file:///C:/VBA/SGQ%201.65/vba-files/Module/modSGQInterface.bas)
- Refine `SetupDashboard`:
    - Rename buttons to user-specified labels ('Système de Gestion', 'Activités de suivi').
    - Shrink the 'Settings' gear icon.
    - Add a global polish routine to iterate through all sheets and apply `DisplayGridlines = False` and `Zoom = 100`.

### [Build & Sync]

#### [EXECUTE] [Global Import](pwsh -File C:\VBA\SGQ%201.65\scripts\Start-VbaAction.ps1 -Action Import)
- Relaunch the full synchronization of modules into the workbook.

---

## Verification Plan

### Automated Tests
- Run compilation check:
  ```powershell
  pwsh -File C:\VBA\SGQ%201.65\scripts\Start-VbaAction.ps1 -Action Compile
  ```

### Manual Verification
1. Open the workbook and verify that the 'Accueil' sheet is the first thing visible.
2. Check that the buttons 'Système de Gestion' and 'Activités de suivi' navigate to 'TDM' and 'TDM-Suivi' respectively.
3. Verify that the 'Settings' gear icon is smaller and functional.
4. Browse other sheets to confirm gridlines are hidden and zoom is at 100%.
