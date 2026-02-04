¾NAttribute VB_Name = "modWorkbookHandlers"
Option Explicit
''AB 20251001_1220
' AUDIT : sauvegarde 20250926_155502 -> vba-files/backups/20250926_155502/modWorkbookHandlers.bas.bak

' =====================================================================================
' Module      : modWorkbookHandlers
' Objectif    : Gestion centralisï¿½e des ï¿½vï¿½nements du classeur SGQ.
' Auteur      : Abel Boudreau, CPA
' Date        : 2025-09-26
' Version     : 2.0.1
'
' CHANGELOG:
'   - 2.0.1 (2025-09-26 15:55:02): ï¿½tape 11.4 - Remplacement des "On Error Resume Next" par des helpers dï¿½diï¿½s.
'   - 2.0.0 (2025-09-24): Refactorisation majeure.
'     - Suppression de la gestion d'ï¿½tat manuelle (modSGQSettings, modSGQAppScopeHelpers).
'     - Standardisation complï¿½te sur CAppStateScope.
'     - Suppression de la boucle d'attente active dans HandleWorkbookActivate.
'     - Nettoyage gï¿½nï¿½ral du formatage.
' =====================================================================================
' ==============================================================================
' Procï¿½dure : HandleWorkbookOpen
' Objet     : ï¿½ l'ouverture : optimise, exï¿½cute les macros contextuelles,
'             puis restaure l'ï¿½tat de l'application via appScope.
' ==============================================================================
Public Sub HandleWorkbookOpen()
Attribute HandleWorkbookOpen.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("HandleWorkbookOpen")
    On Error GoTo ErrorHandler

    ' Dï¿½termination des macros ï¿½ exï¿½cuter selon le nom du fichier
    Dim macroNames As Variant, macroName As Variant, nomFichier As String
    Dim dotPos As Long

    dotPos = VBA.InStrRev(ThisWorkbook.name, ".")
    If dotPos > 0 Then
        nomFichier = VBA.Left$(ThisWorkbook.name, dotPos - 1)
    Else
        nomFichier = ThisWorkbook.name
    End If

    Select Case True
        Case Left$(nomFichier, 5) = "0-SGQ": macroNames = Array("QmgDeactivateAdminMode", "EnableQmgFreezePanesOnMultipleSheets")
        Case Left$(nomFichier, 7) = "6-Suivi": macroNames = Array("QmgDeactivateAdminMode", "EnableQmgFreezePanesOnMultipleSheets2")
        Case Left$(nomFichier, 11) = "SGQ complet": macroNames = Array("QmgActivateAdminMode", "EnableQmgFreezePanesOnMultipleSheets", "EnableQmgFreezePanesOnMultipleSheets2")
        Case Else: macroNames = Array()
    End Select

    ' Exï¿½cution sï¿½curisï¿½e des macros contextuelles
    For Each macroName In macroNames
        Select Case CStr(macroName)
            Case "QmgDeactivateAdminMode": QmgDeactivateAdminMode
            Case "EnableQmgFreezePanesOnMultipleSheets": EnableQmgFreezePanesOnMultipleSheets
            Case "EnableQmgFreezePanesOnMultipleSheets2": EnableQmgFreezePanesOnMultipleSheets2
            Case "QmgActivateAdminMode": QmgActivateAdminMode
        End Select
    Next macroName

    UpdateInterfaceView
    
    ' AmÃ©liorations UX (Step 25)
    SetupDashboard
    ApplyVisualPolish

    ' ï¿½tape 11.4 : ajout sï¿½curisï¿½ du menu contextuel SGQ.
    Dim ctxErr As String
    If IsQmgAdminMode() Then
        If Not TryAddQmgContextMenu(ctxErr) Then
            If LenB(ctxErr) = 0 Then ctxErr = "Impossible d'ajouter le menu contextuel SGQ."
            LogError "modWorkbookHandlers", "HandleWorkbookOpen.TryAddQmgContextMenu", 0, ctxErr
        End If
    End If

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
ErrorHandler:
    LogError "modWorkbookHandlers", "HandleWorkbookOpen", Err.Number, Err.Description
    Resume CleanExit
End Sub

' ==============================================================================
' Procï¿½dure : HandleWorkbookActivate
' Objet     : Quand le classeur reprend le focus : rafraï¿½chit l'UI.
' ==============================================================================
Public Sub HandleWorkbookActivate()
Attribute HandleWorkbookActivate.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("HandleWorkbookActivate")
    On Error GoTo ErrorHandler

    UpdateInterfaceView

    ' Invalidation du ruban si prï¿½sent, sans attente active
    Dim ribbonUi As IRibbonUI
    Dim ribbonErr As String
    ' ï¿½tape 11.4 : rï¿½cupï¿½ration sï¿½curisï¿½e du ruban SGQ.
    If TryGetRibbonReference(ribbonUi, ribbonErr) Then
        If Not ribbonUi Is Nothing Then ribbonUi.InvalidateControl "groupActions"
    ElseIf LenB(ribbonErr) > 0 Then
        LogError "modWorkbookHandlers", "HandleWorkbookActivate.TryGetRibbonReference", 0, ribbonErr
    End If

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
ErrorHandler:
    LogError "modWorkbookHandlers", "HandleWorkbookActivate", Err.Number, Err.Description
    Resume CleanExit
End Sub

' ==============================================================================
' Procï¿½dure : HandleWorkbookDeactivate
' Objet     : Quand le classeur perd le focus : rafraï¿½chit l'UI.
' ==============================================================================
Public Sub HandleWorkbookDeactivate()
Attribute HandleWorkbookDeactivate.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("HandleWorkbookDeactivate")
    On Error GoTo ErrorHandler

    UpdateInterfaceView

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
ErrorHandler:
    LogError "modWorkbookHandlers", "HandleWorkbookDeactivate", Err.Number, Err.Description
    Resume CleanExit
End Sub

' ==============================================================================
' Procï¿½dure : HandleWorkbookBeforeSave
' Objet     : Validation et mise ï¿½ jour de la version (cellule Data!B7).
' ==============================================================================
Public Sub HandleWorkbookBeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
Attribute HandleWorkbookBeforeSave.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("HandleWorkbookBeforeSave")
    On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationAutomatic

    ' Vï¿½rifie l'extension
    Dim nomFichier As String, ext As String
    nomFichier = ThisWorkbook.name
    ext = LCase$(Right$(nomFichier, 5))
    If ext <> ".xlsm" And ext <> ".xlsb" Then
        Cancel = True
        GoTo CleanExit
    End If

    ' Met ï¿½ jour Data!B7 si la feuille et la cellule sont accessibles
    Dim ws As Worksheet, celluleCible As Range
    If SheetExists("Data") Then
        Set ws = ThisWorkbook.sheets("Data")
        Set celluleCible = ws.Range("B7")
        If celluleCible.Locked And ws.ProtectContents Then GoTo CleanExit
    Else
        GoTo CleanExit
    End If

    ' Extraction de la version depuis le nom de fichier via RegExp
    Dim nomFichierNettoye As String, version As String
    nomFichierNettoye = Replace(Replace(Application.WorksheetFunction.Clean(nomFichier), Chr$(160), " "), vbTab, " ")

    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.pattern = "v\s*(\d+(?:\.\d+){1,3})"
    rx.IgnoreCase = True

    If rx.Test(nomFichierNettoye) Then
        Dim m As Object: Set m = rx.Execute(nomFichierNettoye)
        version = m(0).SubMatches(0)
        If Trim$(celluleCible.value) <> version Then celluleCible.value = version
    End If

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
ErrorHandler:
    LogError "modWorkbookHandlers", "HandleWorkbookBeforeSave", Err.Number, Err.Description
    Cancel = True
    Resume CleanExit
End Sub

' ==============================================================================
' Procï¿½dure : HandleWorkbookBeforeClose
' Objet     : Avant fermeture : retire le menu contextuel et restaure l'UI.
' ==============================================================================
Public Sub HandleWorkbookBeforeClose(ByRef Cancel As Boolean)
Attribute HandleWorkbookBeforeClose.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("HandleWorkbookBeforeClose")
    On Error GoTo ErrorHandler

    RemoveQmgContextMenu
    Application.DisplayFullScreen = False

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
ErrorHandler:
    LogError "modWorkbookHandlers", "HandleWorkbookBeforeClose", Err.Number, Err.Description
    Cancel = False
    Resume CleanExit
End Sub


'---------------------------------------------------------------------------------------
' Helper : TryAddQmgContextMenu
' Objectif : Ajoute le menu contextuel SGQ en capturant les erreurs locales.
'---------------------------------------------------------------------------------------
Private Function TryAddQmgContextMenu(ByRef outError As String) As Boolean
    On Error GoTo Handler

    outError = vbNullString
    AddQmgContextMenu
    TryAddQmgContextMenu = True
    Exit Function

Handler:
    outError = Err.Description
    TryAddQmgContextMenu = False
    Err.Clear
End Function

'---------------------------------------------------------------------------------------
' Helper : TryGetRibbonReference
' Objectif : Rï¿½cupï¿½re la rï¿½fï¿½rence du ruban SGQ sans masquer les erreurs globalement.
'---------------------------------------------------------------------------------------
Private Function TryGetRibbonReference(ByRef outRibbon As IRibbonUI, ByRef outError As String) As Boolean
    On Error GoTo Handler

    outError = vbNullString
    Set outRibbon = modRibbonSGQ.GetSGQRibbon
    TryGetRibbonReference = True
    Exit Function

Handler:
    outError = Err.Description
    Set outRibbon = Nothing
    TryGetRibbonReference = False
    Err.Clear
End Function


¾N"(9e377b08b88e30057baf56c074b3d12f1d2237232Bfile:///c:/VBA/SGQ%201.65/vba-files/Module/modWorkbookHandlers.bas:file:///c:/VBA/SGQ%201.65