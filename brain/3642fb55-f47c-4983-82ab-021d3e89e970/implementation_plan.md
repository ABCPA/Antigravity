# Implementation Plan - Secure Update Mechanism

## Goal Description
Implement a "Cloud-first" update mechanism (`modSecureUpdate`) that allows the Excel Add-in to self-update by downloading the latest VBA modules from a centralized SharePoint location via Microsoft Graph.

## Best Practice Workflow: GitHub + SharePoint
Since you use **GitHub**, it should remain your "Source of Truth" (versioning, history).
**SharePoint** will act as your "Distribution Server" (CDN).

**The Flow:**
1.  **Develop**: Commit/Push changes to GitHub.
2.  **Release**: When a version is stable, you "Release" it by copying the files to the SharePoint folder.
3.  **Update**: Users' Excel apps check SharePoint for `manifest.json`.

## User Review Required
> [!IMPORTANT]
> **SharePoint Path Configuration**: The script will be configured to look in the default Document Library for the path: `/General/SGQ_Updates/Production`.
> **Action**: You will need to create this folder structure in your SharePoint "Documents".

## Proposed Changes

### VBA Core
#### [NEW] [modSecureUpdate.bas](file:///C:/VBA/SGQ%201.65/modules/modSecureUpdate.bas)
- **Dependencies**: `modGraphClient` (Updated to use `MSXML2.ServerXMLHTTP` to avoid Access Denied errors), `modConstants`, `JsonConverter`.
- **Functions**:
    - `CheckForUpdates(force As Boolean)`: 
        - Queries Graph: `GET /sites/root/drive/root:/General/SGQ_Updates/Production/manifest.json`.
        - Downloads content and compares versions.
    - `DownloadUpdates(manifest As Object)`:
        - Iterates through `manifest.modules`.
        - Downloads each file from `.../Production/modules/[filename]`.
        - Saves to `tmp/updates/`.
    - `ApplyUpdates()`:
        - Triggers the legacy import logic.

#### [MODIFY] [modConstants.bas](file:///C:/VBA/SGQ%201.65/modules/modConstants.bas)
- Add `GRAPH_UPDATE_ROOT_PATH` = `"/General/SGQ_Updates/Production"`.

## Verification Plan

### Automated Tests
- null

### Manual Verification
- **Setup**: Create folder `/General/SGQ_Updates/Production` in SharePoint. Upload a dummy `manifest.json`.
- **Run**: Execute `modSecureUpdate.TestCheckForUpdates()`.
