# Implementation Plan: Fix Global VBA Compilation Error

## Goal
Resolve the persistent compilation error in `modGraphClient` (undefined `GRAPH_BASE_URL`) which prevents the VBA project from compiling, despite the constant being confirmed present in `modConstants`.

## Diagnosis
- **Symptoms:** `modGraphClient` fails to compile on line 50 (`GRAPH_BASE_URL` not found).
- **Verification:** `modConstants` contains the definition. Re-importing individual modules and in-memory editing failed to resolve it.
- **Root Cause:** High probability of **corrupted VBA Project state** (P-Code desynchronization) or "Zombie Project" syndrome.
- **Solution:** Perform a "Nuclear Rebuild" of the VBA project components.

## Proposed Strategy: Clean Rebuild

We will use a 3-step PowerShell automation to rebuild the project structure from source files.

### 1. Export & Backup
- Script: `Export-AllModules.ps1`
- Action: Export all current 106 components from `index.xlsm` to `c:\VBA\SGQ 1.65\vba-files\clean_source`.
- Validation: Ensure file count matches component count.

### 2. Wipe Project
- Script: `Wipe-VBAProject.ps1`
- Action: Remove ALL components from `index.xlsm` (except `ThisWorkbook` and Sheet modules which can only be cleared, not removed).
- Action: Force save and standard Excel "Open/Close" cycle to flush internal caches.

### 3. Clean Import
- Script: `Import-clean-source.ps1`
- Action: Import all files from `vba-files\clean_source` back into `index.xlsm`.
- Validation: Run `test-vbide-and-compile.ps1` to confirm successful compilation.

## Verification Plan

### Automated Verification
Run the standard compilation check script:
```powershell
pwsh -File "vba-files\scripts\test-vbide-and-compile.ps1" -AttemptCompile -WorkbookPath "index.xlsm"
```
**Success Criteria:**
- `Compile attempted: True`
- `Compile clean: True`
- Exit Code: 0

### Manual Verification
- Open Excel.
- Verify `Debug > Compile` is grayed out (already compiled).
- Verify values of constants in Immediate Window (e.g., `?modConstants.GRAPH_BASE_URL`).

## Risks
- **Sheet Code Loss:** Code in Sheet modules (e.g., `Sheet1`) cannot be "removed" and re-imported easily.
    - *Mitigation:* The export script will save them as `.cls`. The import script must read their content and inject it back into the existing Sheet components instead of trying to add new ones.
- **Reference Loss:** References to external libraries might be reset.
    - *Mitigation:* The Diagnostics script already listed active references. We will verify them post-wipe.

## Files to Create/Modify
- `vba-files/scripts/Rebuild-VBAProject.ps1` (New comprehensive script orchestrating the steps)
