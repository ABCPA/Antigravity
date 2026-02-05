# Implementation Plan - Fix MCP Verification & Compile Check

## Goal Description
Fix false negatives in environment verification scripts. 
1. `verify-mcp-health.ps1` fails to detect running containers because it searches for `docker-mcp-` prefix, while containers are named `mcp-`. It also doesn't return a failure exit code when containers are missing.
2. `test-vbide-and-compile.ps1` returns exit code 0 but implies ambiguity when the project is already compiled (Clean). We need to make it explicitly report "Clean" and ensure the exit code reflects this positive state clearly.

## User Review Required
> [!NOTE]
> None. These are bug fixes to scripts.

## Proposed Changes

### Scripts

#### [MODIFY] [verify-mcp-health.ps1](file:///c:/VBA/SGQ%201.65/scripts/verify-mcp-health.ps1)
- Update regex to match `mcp-` or `docker-mcp-` to be safe: `(?i)(docker-)?mcp-`.
- Update logic to `exit 1` if no containers are found, so the pipeline stops.

#### [MODIFY] [test-vbide-and-compile.ps1](file:///c:/VBA/SGQ%201.65/vba-files/scripts/test-vbide-and-compile.ps1)
- When "Compile" control is disabled (project clean), explicitly set `$global:__CompileClean = $true` so the final exit code logic works correctly (currently `$null`).

## Verification Plan

### Automated Tests
- Run `scripts/verify-mcp-health.ps1` and expect "Conteneurs MCP détectés" and Exit Code 0.
- Run `vba-files/scripts/test-vbide-and-compile.ps1 -AttemptCompile` and expect "Compile clean: True" and Exit Code 0.
