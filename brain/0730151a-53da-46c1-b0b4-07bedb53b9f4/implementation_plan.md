# Implementation Plan - Environment Validation & Script Fixes

Fix the `test-vbide-and-compile.ps1` script and launch the MCP environment to ensure "Audit-grade" reliability.

## Proposed Changes

### VBA Scripts Layer

#### [MODIFY] [test-vbide-and-compile.ps1](file:///c:/VBA/SGQ%201.65/vba-files/scripts/test-vbide-and-compile.ps1)
- Initialize `$global:__CompileClean = $null` at the very beginning of the parameter block or script body to ensure it's always defined.
- Add additional guards for `CommandBars` access to prevent "VBE object is null" or property access errors.
- Ensure the script returns a clear status for both VBIDE access and compile result.

### Docker MCP Layer

#### [EXECUTE] [launch_env.ps1](file:///c:/VBA/SGQ%201.65/docker-mcp/launch_env.ps1)
- Run the script to start the MCP environment.

## Verification Plan

### Automated Tests
- Run `test-vbide-and-compile.ps1` with `-AttemptCompile` flag.
  - Command: `pwsh -File "c:\VBA\SGQ 1.65\vba-files\scripts\test-vbide-and-compile.ps1" -AttemptCompile`
- Run `verify-mcp-health.ps1` to check Docker containers.
  - Command: `pwsh -File "c:\VBA\SGQ 1.65\scripts\verify-mcp-health.ps1"`

### Manual Verification
- Verify that Excel opens and attempts a compilation.
- Confirm that the `mcp-network` and containers are active in Docker Desktop.
