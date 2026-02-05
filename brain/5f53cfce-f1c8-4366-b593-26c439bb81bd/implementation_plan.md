# Implementation Plan - Audit & Cleanup (Phase 27)

The audit revealed that `modSGQUpdateManager.bas` is missing but essential, `modSGQTrackingBuilder.bas` is obsolete due to the "One File" strategy (Phase 24), and `btnCreateSubfolderFile` is disconnected.

## User Review Required

> [!IMPORTANT]
> **Restoration Verification**: I will restore `modSGQUpdateManager.bas` from the backup `export_backup_20251029_105108`. Please confirm this is the correct version source.

> [!WARNING]
> **UI Change**: `btnCreateSubfolderFile` will be reconnected to `CreateClientCopy`. This assumes `CreateClientCopy` is the intended replacement for the old subfolder creation logic.

## Proposed Changes

### Module Restoration

#### [NEW] [modSGQUpdateManager.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQUpdateManager.bas)
- Restore file from `c:\VBA\SGQ 1.65\logs\vba_backups\export_backup_20251029_105108\Module\modSGQUpdateManager.bas`.

### Manifest Update

#### [MODIFY] [manifest.json](file:///c:/VBA/SGQ%201.65/vba-files/manifest.json)
- Add `vba-files/Module/modSGQUpdateManager.bas` to imports.

### UI Fixes

#### [MODIFY] [modRibbonSGQ.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modRibbonSGQ.bas)
- Update `btnCreateSubfolderFile_wrapper` to call `ExecuteActionSafely "CreateClientCopy"`.

### Documentation

#### [MODIFY] [ARCHITECTURE_ET_PLAN.md](file:///c:/VBA/SGQ%201.65/docs/ARCHITECTURE_ET_PLAN.md)
- Update Phase 27 checklist.
- Explicitly note `modSGQTrackingBuilder` as obsolete/removed in Phase 24 notes if not already clear.

## Verification Plan

### Automated Tests
- **Compilation Check**: Run `scripts/test-vbide-and-compile.ps1` to ensure the restored module compiles and integrates correctly.
- **Manifest Validation**: Ensure `manifest.json` is valid JSON.

### Manual Verification
- **Update Check**: User can trigger "Check Update" (if accessible via UI) or inspect the code to see `CheckQmgUpdate` is available.
- **Subfolder Button**: User to test "Créer Sous-dossiers" button to verify it triggers `CreateClientCopy` logic (prompting to create client folder).
