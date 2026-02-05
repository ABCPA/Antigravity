# System Root Audit & Optimization

- [x] **Structural Integrity Check**
    - [x] Validate `workspaces.json` schema and paths
    - [x] Check `.agent` configuration files
    - [x] Inspect `global_workflows` definitions
    - [x] Review `mcp_config.json`

## Audit Findings

### 1. Configuration & Integrity
- **Status**: ✅ **HEALTHY**
- All JSON configurations are valid.
- `global_workflows` contains 9 functional workflows.

### 2. Hygiene Issues
- **Brain**: 62 persistent sessions. No archival policy.
- **Code Tracker**: 8 generic/stale sessions in `active` state.
- **Scratch**: Contains 3 orphaned files, but `autoDiscovery` is `false`.

## Proposed Improvements

### Proposal A: System Hygiene Automations (Recommended)
Create a `sys-cleanup` workflow to:
1. Archive `brain` contexts older than 30 days to `brain/archive`.
2. Prune stale `code_tracker/active` sessions.

### Proposal B: Workflow Optimization
Update `global_workflows` to use `// turbo` annotations for safe, read-only steps (e.g., `gov-check`, `vba-audit-ui`), reducing user clicks.

### Proposal C: Enable Scratch Discovery
Set `autoDiscovery.enabled: true` in `workspaces.json` to officially support ad-hoc work in the `scratch` folder.

- [x] **Content & Hygiene Analysis**
    - [x] Analyze `brain` directory (orphaned flows, size)
    - [x] Check `scratch` and `playground` usage
    - [x] Verify `global_workflows` consistency
- [x] **Functional Improvement Proposals**
    - [x] Draft improvement plan (Architecture, Performance, Features)
    - [x] Review proposals with User
- [x] **Implementation**
    - [x] Enable `autoDiscovery` in `workspaces.json`
    - [x] Create `sys-cleanup` workflow (Brain & Tracker)
    - [x] Optimize global workflows (`// turbo`)
- [x] **Naming Standardization**
    - [x] Rename session artifacts (`SystemRoot_*`)
    - [x] Rename global workflows (`sgq-*`, `sys-*`)
    - [x] Update `workspaces.json` workflow references
- [x] **Verification**
    - [x] Test `sys-cleanup` workflow (Success)
