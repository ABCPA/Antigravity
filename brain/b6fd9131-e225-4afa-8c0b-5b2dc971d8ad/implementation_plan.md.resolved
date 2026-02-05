# Plan: Environment Validation & Path Correction

The goal is to resolve script execution failures related to spaces in the project path and missing Docker dependencies in certain validation scripts.

## Proposed Changes

### [Component] Scripts

#### [MODIFY] [verify-mcp-health.ps1](file:///c:/VBA/SGQ%201.65/scripts/verify-mcp-health.ps1)
- Add logic to detect Docker's absolute path if not in current `env:PATH`.
- Ensure all internal path references are quoted to handle spaces in `c:\VBA\SGQ 1.65`.

#### [MODIFY] [test-vbide-and-compile.ps1](file:///c:/VBA/SGQ%201.65/vba-files/scripts/test-vbide-and-compile.ps1)
- Explicitly quote the `$importWrapper` path when calling `pwsh`.

## Verification Plan

### Automated Tests
1. **Docker Detection**: Run `verify-mcp-health.ps1` and verify it detects Docker even if not in the initial PATH.
2. **VBIDE & Compile**: Run `test-vbide-and-compile.ps1` with the `-AttemptCompile` flag to verify it handles the spaced path correctly when calling the import wrapper.

### Manual Verification
- Confirm that both scripts report success (or specific actionable errors) without failing on "path not found" or "command not recognized".
