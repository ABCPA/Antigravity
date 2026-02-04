# Workspace Audit & Correction Plan

## Issues Identified

### 1. Missing Ribbon Callback
**Problem**: The ribbon XML references `onAction="UpdateInterfaceView"` but `modRibbonSGQ.bas` does not contain this callback.

**Root Cause**: `UpdateInterfaceView` exists in `modSGQInterface.bas` as a public sub, but ribbon callbacks must be in the ribbon module.

**Solution**: Create a wrapper callback in `modRibbonSGQ.bas`:
```vb
Public Sub UpdateInterfaceView(control As IRibbonControl)
    modSGQInterface.UpdateInterfaceView
End Sub
```

---

### 2. Orphaned Brain Folder
**Problem**: `CPA_Unified.code-workspace` referenced an old brain folder (`38c0952b...`).

**Status**: ✅ **FIXED** - Updated to consolidated folder `895b847e-d398-440d-85ff-2aea1b2035f2`.

---

### 3. Broken JSON in Workspace_CPA_AI/mcp.json
**Problem**: File contains trailing garbage after the closing brace (lines 20-23).

**Solution**: Clean up the file to valid JSON.

---

## Proposed Changes

### VBA

#### [MODIFY] [modRibbonSGQ.bas](file:///C:/VBA/SGQ%201.65/vba-files/Module/modRibbonSGQ.bas)

Add the missing ribbon callback after line 680 (after `CallbackToggleView`):

```vb
'---------------------------------------------------------------------------------------
' Callback : UpdateInterfaceView
' Objectif : Callback ruban pour recharger l'interface utilisateur.
'---------------------------------------------------------------------------------------
Public Sub UpdateInterfaceView(control As IRibbonControl)
    modSGQInterface.UpdateInterfaceView
End Sub
```

---

### Configuration

#### [MODIFY] [mcp.json](file:///C:/Users/AbelBoudreau/Workspace_CPA_AI/mcp.json)

Remove trailing invalid JSON (lines 20-23) and ensure proper closure.

---

## Verification Plan

### Automated Tests
- Run VBA compilation: `scripts/test-vbide-and-compile.ps1`
- Validate JSON: `Test-Json (Get-Content mcp.json -Raw)`

### Manual Verification
- Open Excel workbook and verify "Recharger" button works without errors
- Confirm workspace loads correctly in VS Code
