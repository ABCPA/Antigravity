# Implementation Plan - Fix QA Duplicates

## Goal Description
Resolve critical "Ambiguous name detected" errors identified by the QA Agent. This involves refactoring duplicate procedures to use a single source of truth.

## User Review Required
> [!IMPORTANT]
> This refactoring will modify 5+ modules.
> `TrySetSheetVisibility` will be centralized in `modSGQProtection`.

## Proposed Changes

### Administration & Protection
#### [MODIFY] [modSGQAdministration.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQAdministration.bas)
- **Delete**: `TrySetSheetVisibility` and `SetAdminSheetsVisibility`.
- **Modify**: Calls to use `modSGQProtection.TrySetSheetVisibility` (if Public) or import the logic properly.

#### [MODIFY] [modSGQProtection.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQProtection.bas)
- **Ensure**: `TrySetSheetVisibility` is `Public`.

### File System
#### [MODIFY] [modSGQAuditLog.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQAuditLog.bas)
- **Delete**: Private `FolderExists`.
- **Modify**: Use `modSGQFileSystem.FolderExists`.

### Inspection & CQ Rules
#### [MODIFY] [modInspectionAudit.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modInspectionAudit.bas)
- **Rename**: `CreateIssue` -> `CreateInspectionIssue` to disambiguate.

## Verification Plan
### Automated Tests
- Run `scripts/vba-analyze.ps1`.
- Verified success = `Found 0 duplicate(s)` (or at least reduction by 4).
