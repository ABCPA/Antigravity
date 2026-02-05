þ---
description: Standardized Code Review per User Rules
---

1. Ask the user for the review **COMMAND** (e.g., `SYNTAX`, `VARIABLES`, `HEADERS`, `ERRORS`, `ALL`) and **SCOPE** (e.g., `CURRENT`, `ALL`, `[path]`).
2. Read the review prompts from `c:\Users\AbelBoudreau\Workspace_CPA_AI\.vscode\REVIEW_PROMPTS.md`.
3. If SCOPE is `CURRENT`:
   - Identify the currently active VBA file (.bas, .cls, .frm).
   - If no file is active, ask the user to open one.
4. If SCOPE is `ALL`:
   - List all .bas and .cls files in `vba-files\` (excluding backups).
   - Select a subset to review or ask for specific focus if too many.
5. executing the review:
   - Analyze the code strictly against the criteria in `REVIEW_PROMPTS.md`.
   - Check also against `docs\VBA_FILE_HEADER_CONVENTIONS.md` for headers.
   - Output the findings in a structured format (Command: [Type], Scope: [Files]).
þ"(9e377b08b88e30057baf56c074b3d12f1d22372324file:///c:/VBA/SGQ%201.65/.agent/workflows/review.md:file:///c:/VBA/SGQ%201.65