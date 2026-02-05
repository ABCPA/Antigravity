# Phase 23: Secure Update Mechanism

## Goal Description
Implement a robust mechanism to update VBA modules from a JSON manifest/source folder, with automatic backup and rollback capabilities. This addresses Backlog Item 2.

## User Review Required
- **Process**: The update will close/reopen the workbook. Is this acceptable? (Validation required later).

## Proposed Changes

### Core Logic
#### [MODIFY] [vba-files/Module/modSGQUpdateManager.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQUpdateManager.bas)
-   Implement `TryUpdateFromFolder(folderPath)`:
    1.  Validate folder exists.
    2.  Read `manifest.json` from folder.
    3.  **Backup** current modules to `backups/pre_update_YYYYMMDD`.
    4.  Iterate manifest:
        -   Import new modules.
        -   Remove obsolete modules (if specified).
    5.  Log results.

#### [MODIFY] [vba-files/Module/modSGQFileSystem.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQFileSystem.bas)
-   Ensure backup utilities are robust (`BackupCurrentVersion`).

## Verification Plan

### Manual Verification
1.  Create a "dummy" update folder with a modified module.
2.  Run `UpdateModules` macro.
3.  Verify:
    -   Backup created.
    -   Module updated.
    -   Audit log written.
