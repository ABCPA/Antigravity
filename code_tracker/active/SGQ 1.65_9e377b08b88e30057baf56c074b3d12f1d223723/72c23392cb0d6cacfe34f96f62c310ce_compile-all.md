µ---
description: Compile and Verify Entire Project Integrity
---

1. **Context**: Ensure the entire VBA project compiles without errors. This is functionally a full build verification.
2. Run the compilation verification script:
// turbo
3. Run command: `powershell -ExecutionPolicy Bypass -File "c:\VBA\SGQ 1.65\vba-files\scripts\test-vbide-and-compile.ps1"`
4. **Verification**:
   - Confirm output says "COMPILATION OK".
   - If errors exist, they indicate a project-wide consistency issue (broken references, missing dependencies, or syntax errors).
µ"(9e377b08b88e30057baf56c074b3d12f1d22372329file:///c:/VBA/SGQ%201.65/.agent/workflows/compile-all.md:file:///c:/VBA/SGQ%201.65