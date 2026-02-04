# Implementation Plan - UX Improvements

## Goal Description
Enhance the user experience of the SGQ application by introducing a professional "Landing Page" (Dashboard) and applying global visual polish.

## User Review Required
> [!NOTE]
> **Constraint**: We will NOT rename the actual worksheets (e.g. 'TDM' to '💼 TDM') as this would break existing code references. The icons will be prominent on the Dashboard buttons.

## Proposed Changes

### UI/Interface Layer

#### [MODIFY] [modSGQInterface.bas](file:///c%3A/VBA/SGQ%201.65/vba-files/Module/modSGQInterface.bas)
- **New Procedure `SetupDashboard`**:
    - Check if 'Accueil' sheet exists, create if not as the very first sheet.
    - Clear existing content/shapes.
    - Set styling: No gridlines, background color (RGB(245, 247, 250) - soft gray), Zoom 100%.
    - **Draw Buttons**:
        - **Button 1 (SGQ)**:
            - **Icon**: 💼 (Briefcase)
            - **Label**: "Système de Gestion"
            - **Action**: Go to Sheet 'TDM'
            - **Color**: RGB(52, 73, 94) (Dark Slate).
        - **Button 2 (Suivi)**:
            - **Icon**: 🖥️ (Desktop Monitor)
            - **Label**: "Suivi & Monitoring"
            - **Action**: Go to Sheet 'TDM-Suivi'
            - **Color**: RGB(22, 160, 133) (Teal).
        - **Settings**: Gear Icon (Shape) -> Links to settings.
- **New Procedure `ApplyVisualPolish`**:
    - Iterate through all worksheets.
    - set `ActiveWindow.DisplayGridlines = False`.
    - set `ActiveWindow.Zoom = 100`.
    - set `Range("A1").Select`.

## Verification Plan

### Manual Verification
- Code review of the VBA logic.
- User checks after re-import.
