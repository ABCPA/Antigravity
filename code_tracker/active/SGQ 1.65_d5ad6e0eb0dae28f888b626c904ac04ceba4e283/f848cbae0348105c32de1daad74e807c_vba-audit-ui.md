Æ---
description: Audit User Interactions (MsgBox/InputBox) for CPA Professionalism
---

1. Analyze `logs/msgbox_report.csv` if it exists, or run a grep search for "MsgBox" and "InputBox" in `vba-files/`.
2. Identify messages that:
   - Do not use constants (Hardcoded strings).
   - Use unprofessional tone or formatting.
   - Lack error context.
3. Propose a refactoring plan to move strings to `modConstants.bas` with `Public Const MSG_...`.
Æ*cascade08"(d5ad6e0eb0dae28f888b626c904ac04ceba4e2832:file:///c:/VBA/SGQ%201.65/.agent/workflows/vba-audit-ui.md:file:///c:/VBA/SGQ%201.65