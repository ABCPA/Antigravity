# Implementation Plan - ScanCircularDependencies

The goal is to implement the currently stubbed `ScanCircularDependencies` procedure in `modDiagnostics.bas`. This function was identified as "Complete" in the architecture plan but exists only as a placeholder. It aims to detect potential circular calls or dependencies between modules.

## User Review Required

> [!IMPORTANT]
> This implementation relies on **Static Code Analysis** (parsing `.bas` files), not runtime tracking, as VBA does not support reflection. It will scan the exported `vba-files/` directory.

## Proposed Changes

### Core & Utilities Layer

#### [MODIFY] [modDiagnostics.bas](file:///c:/VBA/SGQ%201.65/vba-files/Module/modDiagnostics.bas)
- Replace the `ScanCircularDependencies` stub with actual logic.
- **Logic**:
    1. List all `.bas` and `.cls` files in `vba-files/`.
    2. Parse each file to find `Public Sub`/`Function` definitions.
    3. Parse each file to find calls to these procedures (simple text matching or regex).
    4. Build a directed graph (Adjacency Matrix or List).
    5. Run a cycle detection algorithm (DFS).
- **Output**:
    - Print cycles to the Immediate Window.
    - Log details to `logs/diagnostics/circular_deps_YYYYMMDD.txt` (via `modSGQUtilitaires.LogError` or dedicated logger).

## Verification Plan

### Automated Tests
- Run `ScanCircularDependencies` from the Immediate Window.
- Verify it detects a known cycle (can create a temporary cycle for testing).
- Check `logs/diagnostics` for output.

### Manual Verification
- **Test**: Create `modCycleA` calling `modCycleB`, and `modCycleB` calling `modCycleA`.
- **Run**: `ScanCircularDependencies`.
- **Expect**: The cycle `modCycleA -> modCycleB -> modCycleA` is reported.
