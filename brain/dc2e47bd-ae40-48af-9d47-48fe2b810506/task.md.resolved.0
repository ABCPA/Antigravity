# Task: Résolution des erreurs d'encodage (Mojibake)

## Diagnosis & Repair
- [ ] Update `ARCHITECTURE_ET_PLAN.md` with current objective [/]
- [ ] Identify all files containing mojibake patterns (e.g., "Ã©", "Ã¨", "Ã")
- [ ] Fix encoding in `modSGQUpdateManager.bas`
- [ ] Fix encoding in `modSGQCreation.bas`
- [ ] Fix encoding in `modRibbonSGQ.bas`
- [ ] Fix encoding in `modSGQUIActionDispatcher.bas`
- [ ] Verify no other files are affected
- [ ] Compile and verify project (`test-vbide-and-compile.ps1`)

## Prevention (Tooling)
- [ ] Audit PowerShell import/export scripts for encoding enforcement (CP1252 vs UTF-8)
- [ ] Create or update a script to normalize encodings to CP1252 automatically
- [ ] Update documentation on file encoding standards

## Verification
- [ ] Manual check of critical UI messages (MsgBox)
- [ ] Final Compilation check
