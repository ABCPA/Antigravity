# Implementation Plan: Script Fixes and Environment Launch

Address script failures in the validation pipeline and ensure the MCP environment is operational.

## Proposed Changes

### PowerShell Scripts

#### [MODIFY] [ExcelSafeOpen.ps1](file:///c:/VBA/SGQ%201.65/vba-files/scripts/ExcelSafeOpen.ps1)
- Remove duplicate function definitions.
- Ensure the most robust version of `New-ExcelApplicationSafe` and `Restore-ExcelApplicationState` is kept.

#### [MODIFY] [test-vbide-and-compile.ps1](file:///c:/VBA/SGQ%201.65/vba-files/scripts/test-vbide-and-compile.ps1)
- Initialize `$global:__CompileClean = $null` at the start of the script.
- Wrap `VBE.CommandBars` access in a more robust check.
- Ensure the exit code logic handles the `$null` state of `$global:__CompileClean`.

## Verification Plan

### Automated Tests
1. **Launch MCP Environment**:
   ```powershell
   & "c:\VBA\SGQ 1.65\docker-mcp\launch_env.ps1"
   ```
2. **Verify MCP Health**:
   ```powershell
   & "c:\VBA\SGQ 1.65\scripts\verify-mcp-health.ps1"
   ```
3. **Run Compile Test**:
   ```powershell
   & "c:\VBA\SGQ%201.65\vba-files\scripts\test-vbide-and-compile.ps1" -AttemptCompile
   ```

### Manual Verification
- Check the output of `docker ps` to ensure all containers (Memory, Filesystem, Fetcher, etc.) are running.
- Confirm that the Excel workbook opens and closes without manual intervention during the compile test.
