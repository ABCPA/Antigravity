# Implementation Plan - Fix Encoding (Mojibake)

## Goal
Resolve encoding issues where accented characters appear as "Ã©" (Mojibake) in VBA files. This is caused by UTF-8 byte sequences being interpreted as Windows-1252 characters. The goal is to sanitize the `.bas` and `.cls` files so they are pure Windows-1252 and display correctly in the VBA Editor.

## User Review Required
> [!IMPORTANT]
> This process involves a bulk search-and-replace on the codebase. While backed by git, a backup of the `vba-files` directory is recommended before running the script.

## Proposed Changes

### Scripts
#### [NEW] [scripts/fix-mojibake.ps1](file:///c:/VBA/SGQ%201.65/scripts/fix-mojibake.ps1)
- A PowerShell script to scan all `.bas` and `.cls` files in `vba-files/`.
- It will perform text replacements for known UTF-8 -> Windows-1252 artifacts:
    - `Ã©` -> `é`
    - `Ã¨` -> `è`
    - `Ã ` -> `à` (handling NBSP if necessary)
    - `Ãª` -> `ê`
    - `Ã´` -> `ô`
    - `Ã§` -> `ç`
    - `Ã¹` -> `ù`
    - `Â«` -> `«`
    - `Â»` -> `»`
    - `â€™` -> `’` (Windows-1252 curly quote)
    - And other common patterns.
- The script will save the files back with **Windows-1252** encoding to ensure consistency.

### Tooling
- Update `scripts/import-vba-release.ps1` (if applicable) to enforce Content-Encoding when reading/writing to prevent re-introduction of UTF-8.

## Verification Plan

### Automated Verification
- **Grep Search**: Run `grep_search` for the pattern `Ã©` after the fix. It should return 0 results.
- **Compilation**: Run `scripts/test-vbide-and-compile.ps1` to ensure the code still compiles (no syntax errors introduced).

### Manual Verification
- **Code Review**: Open `modSGQCreation.bas` and `modSGQUpdateManager.bas` in VS Code (reopened with correct encoding) or Notepad to verify accents look correct.
- **Excel Check**: Import the modules and check `MsgBox` outputs in the Excel application.
