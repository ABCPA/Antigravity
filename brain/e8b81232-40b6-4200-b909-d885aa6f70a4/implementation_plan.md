# Plan: Commit Strategy for Recent Changes

The current workspace contains a significant number of changes across different domains: virtual agents, diagnostic scripts, and VBA core refactoring. This plan outlines the strategy to commit these changes professionally.

## Proposed Strategy

I propose to split the changes into several logical commits to maintain a clean and understandable history:

1.  **feat(agents): add virtual agent infrastructure and workflows**
    - Includes `.agent/agents/` and `.agent/workflows/`.
2.  **feat(diag): add comprehensive VBA diagnostic scripts**
    - Includes `scripts/` and `vba-files/scripts/`.
3.  **refactor(core): apply architectural standards and fix encoding**
    - Includes modified modules in `vba-files/Module/`.
    - Focus: CP1252 encoding and `appScope` integration.

## Proposed Changes

### [Git Management]

#### [COMMIT] feat(agents): add virtual agent infrastructure
- `.agent/agents/git-agent.md`
- `.agent/agents/diagnostic-agent.md`
- `.agent/workflows/invoke-agent.md`
- ... other agent/workflow files

#### [COMMIT] feat(diag): add advanced VBA diagnostic and repair scripts
- `vba-files/scripts/update_diagnostics.ps1`
- `vba-files/scripts/Diagnose-VBACompilation.ps1`
- ... other script files

#### [COMMIT] refactor(core): standardize modules for CP1252 and appScope
- `vba-files/Module/modSGQInterface.bas`
- `vba-files/Module/modConstants.bas`
- ... other modified modules

## Verification Plan

### Automated Tests
- `scripts\test-vbide-and-compile.ps1` (Validation de l'intégrité du projet VBA)

### Manual Verification
- Review resulting `git log` to ensure commits are atomic and descriptive.
