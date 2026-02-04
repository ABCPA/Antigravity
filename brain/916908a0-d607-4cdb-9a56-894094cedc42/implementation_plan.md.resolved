# Implementation Plan - Refactoring and Encoding Fixes

## Goal Description
Refactor `CreateSubfolderFile` to remove duplicate headers and use the correct folder structure.
Clean up VBA modules (`modSGQInterface.bas`, `modSGQViews.bas`, etc.) that are corrupted with ANSI escape codes (e.g., `[7m...[0m`) and mojibake.
Ensure all modules are correctly encoded (CP1252).

## User Review Required
> [!IMPORTANT]
> Several files seem to be corrupted with ANSI escape sequences (likely from a previous `grep --color=always` output pasted into the file). I will strip these sequences.

## Proposed Changes

### Refactoring
#### [MODIFY] [modSGQCreation.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQCreation.bas)
- Remove duplicate procedure header for `CreateSubfolderFile`.
- Ensure `CreateSubfolderFile` uses `App.PathSeparator` and constants.

### Encoding & Corruption Fixes
I will use a PowerShell script to clean the files.

#### [MODIFY] [modSGQInterface.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQInterface.bas)
- Remove ANSI escape codes.
- Fix mojibake (e.g. `Syst?me` -> `Système`).

#### [MODIFY] [modSGQViews.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQViews.bas)
- Remove ANSI escape codes.

#### [MODIFY] [modSGQUtilitaires.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQUtilitaires.bas)
- Check and clean if affected.

## Verification Plan

### Automated Tests
- **Compilation Check**: Run `/compile` (invokes `test-vbide-and-compile.ps1`) to ensure no syntax errors remain.
- **Encoding Check**: Run `scripts/check-attributes-local.ps1` to verify CP1252 encoding.
- **Grep Check**: Search for `\[[0-9;]*m` to ensure all ANSI codes are gone.

### Manual Verification
- Review diffs to ensure no code logic was lost during cleaning.
