# Task: Project Maintenance & Documentation Cleanup

## Status: Planning

- [x] **Documentation Cleanup**
    - [x] Analyze `ARCHITECTURE_ET_PLAN.md` for duplicate sections (Phase 0 repeated). <!-- id: 1 -->
    - [x] Create a consolidated version of `ARCHITECTURE_ET_PLAN.md`. <!-- id: 2 -->
    - [x] Verify `GEMINI.md` for any redundancy. <!-- id: 3 -->
- [x] **Project Health Check**
    - [x] Run `test-vbide-and-compile.ps1` to ensure VBA project is stable. <!-- id: 4 -->
    - [x] Verify `vba-files/manifest.json` integrity. <!-- id: 5 -->
- [x] **Verification**
    - [x] Commit cleanup changes. <!-- id: 6 -->
- [x] **Bug Fix: Ambiguous Name Compilation Error**
    - [x] Diagnose `UpdateInterfaceView` duplication in `modWorkbookHandlers` / `modSGQInterface`. <!-- id: 7 -->
    - [x] Resolve ambiguity (verify source files are clean). <!-- id: 8 -->
- [x] **Design VBA QA Pipeline**
    - [x] Create `qa-agent.md` definition. <!-- id: 9 -->
    - [x] Create `scripts/vba-analyze.ps1` for static analysis report. <!-- id: 10 -->
    - [x] Verify `modVBAInspector.bas` capabilities. <!-- id: 11 -->
    - [x] Document workflow in `ARCHITECTURE_ET_PLAN.md`. <!-- id: 12 -->
- [x] **Generate QA Report**
    - [x] Locate latest JSON report in `logs/qa/`. <!-- id: 13 -->
    - [x] Generate readable Markdown summary of duplicates. <!-- id: 14 -->
- [x] **Fix High Priority QA Duplicates**
    - [x] `TrySetSheetVisibility`: Consolidate to `modSGQProtection`. <!-- id: 15 -->
    - [x] `FolderExists`: Consolidate to `modSGQFileSystem`. <!-- id: 16 -->
    - [x] `CreateIssue`: Clarify usage in `modCQRules` vs `modInspectionAudit`. <!-- id: 17 -->
    - [x] Verify fixes with re-scan. <!-- id: 18 -->
- [x] **Commit and Sync**
    - [x] Stage all modified files. <!-- id: 19 -->
    - [x] Generate conventional commit message. <!-- id: 20 -->
    - [x] Push changes. <!-- id: 21 -->
