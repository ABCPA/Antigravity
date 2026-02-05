¾---
description: Guide to refactor a feature to the 4-layer architecture
---

1. Identify the module(s) to refactor.
2. Read `docs\ARCHITECTURE_ET_PLAN.md` to determine the target layer (UI, Application, Core, Developer).
3. Create a plan for the refactoring:
    - Identify dependencies to centralize (e.g., `LogError`, `appScope`).
    - Standardize error handling using `On Error GoTo Handler`.
    - Ensure `Option Explicit` is present.
4. Execute the changes.
5. Update `vba-files\manifest.json` if modules were renamed or added.
6. Verify via `/compile`.
¾*cascade08"(c22bc49c04c4165d7ce5af31087440360715f2332>file:///c:/VBA/SGQ%201.65/.agent/workflows/refactor-feature.md:file:///c:/VBA/SGQ%201.65