# System Root Audit & Improvement Plan

## Goal
Perform a comprehensive health check of the Antigravity System Root (`.gemini/antigravity`) to ensure operational integrity, identify misconfigurations, and propose architectural improvements.

## User Review Required
> [!NOTE]
> This plan focuses on **analysis and non-destructive cleanup proposals**. No critical files will be deleted without explicit approval.

## Audit Scope

### 1. Configuration Validation
- **`workspaces.json`**: Verify paths exists, JSON validity, schema compliance.
- **`mcp_config.json`**: Check server definitions.
- **`.agent/orchestrator.json`**: Verify prompt paths and settings.

### 2. Workflow Integrity
- **`global_workflows/`**: Ensure all workflows defined in `workspaces.json` exist and are valid markdown.

### 3. Data Hygiene ("Brain" Health)
- **`brain/`**: Analyze accumulation of session data.
- **`code_tracker/`**: Check for stale state.

## Proposed Improvements (Draft)
*To be refined after audit completion.*

1. **Auto-Discovery**: Enable auto-discovery for `scratch` folder.
2. **Workflow Consolidation**: Review if global workflows need updates.
3. **Brain Maintenance**: Propose a cleanup policy for old sessions.

## 4. Naming Convention Standards (New)

### Goal
Ensure all artifacts and workflows are identifiable by filename without opening them.

### Artifacts (Brain)
Rename generic files to include Context and Artifact Type:
- `task.md` -> `[Context]_Task.md`
- `implementation_plan.md` -> `[Context]_Plan.md`
- `walkthrough.md` -> `[Context]_Walkthrough.md`
*Example*: `SystemRoot_Task.md`, `SystemRoot_Plan.md`

### Workflows (Global)
Prefix workflows with their specific application domain:
- **System**: `sys-[action].md` (e.g., `sys-cleanup.md`, `sys-git-commit.md`)
- **SGQ**: `sgq-[action].md` (e.g., `sgq-vba-compile.md`, `sgq-fix-mojibake.md`)
- **Generic**: `[domain]-[action].md`

### Implementation Steps
1. Rename current session artifacts.
2. Rename existing global workflows to match new convention.
3. Update `workspaces.json` or scripts if they reference specific workflow filenames.

## Verification
- Manual verification of JSON configurations.
- Validation script for workflows (if applicable).
