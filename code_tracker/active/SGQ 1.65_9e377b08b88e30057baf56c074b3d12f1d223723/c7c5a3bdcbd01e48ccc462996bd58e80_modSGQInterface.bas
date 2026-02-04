��Attribute VB_Name = "modSGQInterface"
Option Explicit
'AB20251103_1600' AUDIT : sauvegarde 20251216_233619 -> vba-files/backups/20251216_233619/modSGQInterface.bas.bak' AUDIT : sauvegarde 20251023_131232 -> vba-files/backups/20251023_131232/modSGQInterface.bas.bak
' AUDIT : sauvegarde 20250929_100500 -> vba-files/backups/20250929_100500/modSGQInterface.bas.bak
' AUDIT : sauvegarde 20251008_150550 -> vba-files/backups/20251008_150550/modSGQInterface.bas.bak
' AUDIT : sauvegarde 20251217_162000 -> vba-files/backups/20251217_162000/modSGQInterface.bas.bak

' ===============================================================================
' Module      : modSGQInterface
' Objectif    : G�rer les mises � jour de l'interface utilisateur (UI) et les diagnostics.
' Auteur      : Abel Boudreau, CPA
' Date        : 2025-09-29
' Version     : 2.1.0
'
' CHANGELOG:
' - 2.1.0 (2025-10-23): Step 11.4 - Refactor ConfigureLimitScrollArea and SetAdminColumnsHidden.
'     - Add TrySetWorksheetScrollArea and TrySetColumnsHidden helpers.
'     - Replace On Error Resume Next in loops with structured error handling.
'   - 2.0.3 (2025-09-29 10:05:00): MAJ en-t�te (AUDIT, version) sans changement fonctionnel.
'   - 2.0.2 (2025-09-26 15:20:00): Ajout des helpers TryUpdateInterfaceView et TryShowQmgNotification (gestion d'erreur structur�e et fallback MsgBox).
'   - 2.0.1 (2025-09-26 12:07:32): �tape 11.4 - Affichage UI s�curis� (UpdateInterfaceView/ShowQmgNotification).
'   - 3.0.0 (2025-09-24): Refactorisation majeure.
'     - Suppression des d�pendances � modSGQSettings et modSGQAppScopeHelpers.
'     - Standardisation sur CAppStateScope pour la gestion d'�tat.
'     - Centralisation de la logique de masquage de lignes via HideEmptyRowsInRange.
'     - D�placement des helpers de noms d�finis vers modExcelUtils.
' ===============================================================================

'---------------------------------------------------------------------------------------
' Procedure : UpdateInterfaceView
' Objectif  : G�rer la visibilit� de la barre de formule et du ruban Excel en fonction
'             du classeur actif et du mode administrateur.
'---------------------------------------------------------------------------------------

Public Sub UpdateInterfaceView()
Attribute UpdateInterfaceView.VB_ProcData.VB_Invoke_Func = " \n14"
    ' �tape 11.4 : mise � jour UI avec gestion d'erreur structur�e.
    Dim uiErr As String

    If Not TryUpdateInterfaceView(uiErr) Then
        If LenB(uiErr) > 0 Then
            LogError "modSGQInterface", "UpdateInterfaceView", 0, uiErr
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TryUpdateInterfaceView
' Objectif  : Appliquer la vue UI (barre de formule et ruban) selon le contexte SGQ.
' Param�tres: errMsg (ByRef) - message d'erreur d�taill� si �chec.
' Retour    : Boolean - True si r�ussi, False sinon.
' Remarque  : Compatible Excel 2013 � 365 (utilise SHOW.TOOLBAR pour le ruban).
'---------------------------------------------------------------------------------------
Private Function TryUpdateInterfaceView(ByRef errMsg As String) As Boolean
    On Error GoTo ErrHandler
    Dim hideUi As Boolean

    ' Activer le mode plein cran et cacher la barre de formule seulement pour ce classeur et hors mode admin
    hideUi = (Not ActiveWorkbook Is Nothing) And (ActiveWorkbook Is ThisWorkbook) And (Not IsQmgAdminMode())

    If hideUi Then
        Application.DisplayFormulaBar = False
        ' Mode utilisateur : passer en plein ecran et reduire le ruban (onglets visibles)
        On Error Resume Next
        ' S'assurer que le ruban est visible avant de le reduire (certaines versions masquent le ruban)
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",TRUE)"
        ' Reduire le ruban afin que seuls les onglets restent visibles (ExecuteMso est la facon la plus fiable)
        Application.CommandBars.ExecuteMso "MinimizeRibbon"
        On Error GoTo ErrHandler
        ' Application.DisplayFullScreen = True ' DESACTIVE CAR CACHE LES ONGLETS DU BAS
    Else
        ' Restaurer le ruban et la barre de formule
        On Error Resume Next
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",TRUE)"
        On Error GoTo ErrHandler
        Application.DisplayFormulaBar = True
        Application.DisplayFullScreen = False
    End If
    
    ' FORCER L'AFFICHAGE DES ONGLETS DU BAS
    ActiveWindow.DisplayWorkbookTabs = True

    TryUpdateInterfaceView = True
    Exit Function

ErrHandler:
    errMsg = "TryUpdateInterfaceView failed: " & Err.Description
    TryUpdateInterfaceView = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : EnsureRibbonMinimized
' Objectif  : Minimise le ruban (Show Tabs Only) et assure que les onglets sont visibles.
'---------------------------------------------------------------------------------------
Public Sub EnsureRibbonMinimized()
    On Error Resume Next
    If Application.CommandBars.GetPressedMso("MinimizeRibbon") = False Then
        Application.CommandBars.ExecuteMso "MinimizeRibbon"
    End If
    ' S'assurer que les onglets du bas sont visibles
    ActiveWindow.DisplayWorkbookTabs = True
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : QmgShowRibbonFromUserMode
' Objectif  : Restaure l'interface (affiche le ruban et quitte le plein ecran) lorsque
'             l'utilisateur demande explicitement la reapparition du ruban (ex: Ctrl+F1).
'---------------------------------------------------------------------------------------
Public Sub QmgShowRibbonFromUserMode()
    On Error GoTo Handler
    ' Restaure l'UI pour permettre a l'utilisateur d'acceder facilement aux commandes
    Application.ScreenUpdating = True
    Application.DisplayFormulaBar = True
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",TRUE)"
    On Error GoTo Handler
    Application.DisplayFullScreen = False
    ActiveWindow.DisplayWorkbookTabs = True
    Exit Sub
Handler:
    LogError "modSGQInterface", "QmgShowRibbonFromUserMode", Err.Number, Err.Description
    Err.Clear
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowParameters
' Objectif  : Affiche l'�tat actuel de divers param�tres de l'application Excel
'             dans la fen�tre d'ex�cution (Immediate Window).
'---------------------------------------------------------------------------------------
Public Sub ShowParameters()
Attribute ShowParameters.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("ShowParameters")
    On Error GoTo Handler

    Debug.Print String(50, "=")
    Debug.Print "�tat des param�tres Excel (courant)"
    Debug.Print "- Barre Formule : " & Application.DisplayFormulaBar
    Debug.Print "- Barre d'�tat  : " & Application.DisplayStatusBar
    If Not ActiveWindow Is Nothing Then Debug.Print "- En-t�tes      : " & ActiveWindow.DisplayHeadings
    Debug.Print "- Calcul        : " & Application.Calculation
    Debug.Print "- ScreenUpdating: " & Application.ScreenUpdating
    Debug.Print "- EnableEvents  : " & Application.EnableEvents

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQInterface", "ShowParameters", Err.Number, Err.Description
    Resume CleanExit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowQmgInfo
' Objectif  : Affiche une bo�te de message contenant des informations sur le syst�me SGQ.
'---------------------------------------------------------------------------------------
Public Sub ShowQmgInfo()
Attribute ShowQmgInfo.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("ShowQmgInfo")
    On Error GoTo Handler

    MsgBox MSG_INFO_SYSTEM_DEVELOPER & vbCrLf & _
        MSG_INFO_VBA_VERSION & GetCellValue("Data", "B7") & vbCrLf & _
        MSG_INFO_FILE_VERSION & GetCellValue("Entete", "A1") & vbCrLf & _
        MSG_INFO_CABINET & GetCellValue("Data", "B16"), vbInformation, TITLE_ABOUT_SGQ

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQInterface", "ShowQmgInfo", Err.Number, Err.Description
    Resume CleanExit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowQmgNotification
' Objectif  : Affiche une notification � l'utilisateur, en utilisant un formulaire
'             personnalis� si disponible, sinon une bo�te de message standard.
' Param�tres:
'   message: Le message � afficher dans la notification.
'---------------------------------------------------------------------------------------

Public Sub ShowQmgNotification(Optional ByVal message As String = "Op�ration SGQ termin�e avec succ�s.")
Attribute ShowQmgNotification.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Etape 11.4 : notification avec fallback structure.
    Dim notifyErr As String

    If Not TryShowQmgNotification(message, notifyErr) Then
        If LenB(notifyErr) > 0 Then
            LogError "modSGQInterface", "ShowQmgNotification", 0, notifyErr
        End If
        MsgBox message, vbInformation, TITLE_NOTIFICATION
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TryShowQmgNotification
' Objectif  : Afficher une notification via un UserForm si disponible.
' Param�tres: message (In) - texte � afficher; errMsg (ByRef) - erreur si �chec.
' Retour    : Boolean - True si le formulaire a ete affiche, False si fallback requis.
'---------------------------------------------------------------------------------------
Private Function TryShowQmgNotification(ByVal message As String, ByRef errMsg As String) As Boolean
    On Error GoTo ErrHandler

    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmNotificationSGQ")
    frm.lblMessage.Caption = message
    frm.StartUpPosition = 1
    frm.Show vbModeless

    TryShowQmgNotification = True
    Exit Function

ErrHandler:
    errMsg = "TryShowQmgNotification failed: " & Err.Description
    TryShowQmgNotification = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : RunQmgComplianceDiagnostics
' Objectif  : Ex�cute une s�rie de diagnostics de conformit� sur le classeur SGQ,
'             incluant la v�rification de l'acc�s au projet VBA et des macros de boutons.
'---------------------------------------------------------------------------------------
Public Sub RunQmgComplianceDiagnostics()
Attribute RunQmgComplianceDiagnostics.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("RunQmgComplianceDiagnostics")
    On Error GoTo Handler

    Dim reportText As String
    reportText = "DIAGNOSTIC DU CLASSEUR SGQ" & vbCrLf & String(40, "-") & vbCrLf
    reportText = reportText & "1. Acc�s projet VBA : " & IIf(IsVbideTrusted(), "ACTIV�", "D�SACTIV�") & vbCrLf
    reportText = reportText & vbCrLf & "2. V�rification des macros li�es aux boutons..." & vbCrLf

    CheckButtonMacros

    MsgBox reportText, vbInformation, CAPTION_DIAGNOSTIC_SGQ

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQInterface", "RunQmgComplianceDiagnostics", Err.Number, Err.Description
    Resume CleanExit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnableQmgLimitScrollArea
' Objectif  : Configure les zones de d�filement (ScrollArea) pour des feuilles sp�cifiques
'             afin de limiter la navigation de l'utilisateur � des plages d�finies.
'---------------------------------------------------------------------------------------
Public Sub EnableQmgLimitScrollArea()
Attribute EnableQmgLimitScrollArea.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("EnableQmgLimitScrollArea")
    On Error GoTo Handler

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "TDM", "A1:T34"
    dict.Add "2.10-1", "A1:G61"
    dict.Add "2.10-2", "A1:M22"
    dict.Add "6.20.1", "A1:I103"
    dict.Add "5.60-7", "A1:J35"

    ConfigureLimitScrollArea dict

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQInterface", "EnableQmgLimitScrollArea", Err.Number, Err.Description
    Resume CleanExit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ConfigureLimitScrollArea
' Objectif  : Applique une zone de d�filement (ScrollArea) � une ou plusieurs feuilles
'             sp�cifi�es dans un dictionnaire.
' Param�tres:
'   scrollAreas: Un objet Dictionary o� les cl�s sont les noms des feuilles et les
'                valeurs sont les plages de d�filement (ex: "A1:G50").
'---------------------------------------------------------------------------------------
Public Sub ConfigureLimitScrollArea(ByVal scrollAreas As Object)
Attribute ConfigureLimitScrollArea.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("ConfigureLimitScrollArea")
    On Error GoTo Handler

    Dim key As Variant
    Dim ws As Worksheet
    Dim errMsg As String

    If scrollAreas Is Nothing Then Exit Sub

    For Each key In scrollAreas.Keys
        If Not TrySetWorksheetScrollArea(CStr(key), CStr(scrollAreas(key)), errMsg) Then
            LogError "modSGQInterface", "ConfigureLimitScrollArea", 0, _
                     "Failed to set scroll area for " & CStr(key) & ": " & errMsg
        End If
    Next key

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQInterface", "ConfigureLimitScrollArea", Err.Number, Err.Description
    Resume CleanExit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ResetScrollArea
' Objectif  : R�initialise les zones de d�filement (ScrollArea) pour toutes les feuilles
'             du classeur, permettant une navigation libre.
' Param�tres:
'   afficherMessage: (Optionnel) Si True, affiche un message dans la fen�tre d'ex�cution.
'---------------------------------------------------------------------------------------
Public Sub ResetScrollArea(Optional ByVal afficherMessage As Boolean = True)
Attribute ResetScrollArea.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim appScope As Object
    Set appScope = BeginAppStateScope("ResetScrollArea")
    On Error GoTo Handler

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.sheets
        ws.scrollArea = ""
    Next ws
    If afficherMessage Then Debug.Print "Zones de d�filement r�initialis�es."

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQInterface", "ResetScrollArea", Err.Number, Err.Description
    Resume CleanExit
End Sub

' Re-inserted SetAdminColumnsHidden and Show/Hide wrappers (kept below to avoid reordering other procedures)
Private Sub SetAdminColumnsHidden(ByVal targetSheet As Worksheet, ByVal columnRange As String, ByVal hidden As Boolean)
    Dim appScope As Object
    Set appScope = BeginAppStateScope("SetAdminColumnsHidden")
    On Error GoTo ErrHandler

    If targetSheet Is Nothing Then
        Err.Raise vbObjectError + 513, "SetAdminColumnsHidden", "La feuille cible (TargetSheet) ne peut pas �tre Nothing."
    End If

    If Len(Trim$(columnRange)) = 0 Then
        Err.Raise vbObjectError + 514, "SetAdminColumnsHidden", "La plage de colonnes (ColumnRange) ne peut pas �tre vide."
    End If

    Dim parts As Variant, p As Variant
    Dim changedCount As Long: changedCount = 0
    Dim errMsg As String

    parts = Split(columnRange, ",")
    For Each p In parts
        p = Trim$(p)
        If Len(p) > 0 Then
            If TrySetColumnsHidden(targetSheet, CStr(p), hidden, errMsg) Then
                changedCount = changedCount + 1
            Else
                LogError "modSGQInterface", "SetAdminColumnsHidden", 0, _
                         "Failed to set hidden for columns " & CStr(p) & ": " & errMsg
            End If
        End If
    Next p

    If changedCount > 0 Then
        LogError "modSGQInterface", "SetAdminColumnsHidden", 0, CStr(changedCount) & " column blocks processed for " & targetSheet.name & " (" & columnRange & ")"
    End If

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub

ErrHandler:
    LogError "modSGQInterface", "SetAdminColumnsHidden", Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub ShowAdminColumns(ByVal targetSheet As Worksheet, ByVal columnRange As String)
Attribute ShowAdminColumns.VB_ProcData.VB_Invoke_Func = " \n14"
    SetAdminColumnsHidden targetSheet, columnRange, False
End Sub

Public Sub HideAdminColumns(ByVal targetSheet As Worksheet, ByVal columnRange As String)
Attribute HideAdminColumns.VB_ProcData.VB_Invoke_Func = " \n14"
    SetAdminColumnsHidden targetSheet, columnRange, True
End Sub

' ===============================================================================
' Proc�dures de masquage de lignes vides (utilisant le helper centralis�)
' ===============================================================================

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows42
' Objectif  : Masque les lignes vides dans une plage sp�cifique de la feuille "4.30-2".
'             Laisse visible la premi�re ligne vide pour permettre la saisie de donn�es.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows42()
Attribute HideEmptyRows42.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet42"), "C4:C83"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows44
' Objectif  : Masque les lignes vides dans la plage B4:B34 de la feuille correspondant � Sheet44.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows44()
Attribute HideEmptyRows44.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet44"), "B4:B34"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows45
' Objectif  : Masque les lignes vides dans la plage B4:B34 de la feuille correspondant � Sheet45.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows45()
Attribute HideEmptyRows45.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet45"), "B4:B34"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows46
' Objectif  : Masque les lignes vides dans la plage B4:B34 de la feuille correspondant � Sheet46.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows46()
Attribute HideEmptyRows46.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet46"), "B4:B34"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows47
' Objectif  : Masque les lignes vides dans la plage B4:B34 de la feuille correspondant � Sheet47.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows47()
Attribute HideEmptyRows47.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet47"), "B4:B34"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows48
' Objectif  : Masque les lignes vides dans la plage B4:B34 de la feuille correspondant � Sheet48.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows48()
Attribute HideEmptyRows48.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet48"), "B4:B34"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows491
' Objectif  : Masque les lignes vides dans la plage B4:B34 de la feuille correspondant � Sheet491.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows491()
Attribute HideEmptyRows491.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet491"), "B4:B34"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows492
' Objectif  : Masque les lignes vides dans la plage B4:B34 de la feuille correspondant � Sheet492.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows492()
Attribute HideEmptyRows492.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet492"), "B4:B34"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows65
' Objectif  : Masque les lignes vides dans une plage sp�cifique de la feuille "6.30".
'             Laisse visible la premi�re ligne vide pour permettre la saisie de donn�es.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows65()
Attribute HideEmptyRows65.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange "6.30", "B14:G43", -1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows66
' Objectif  : Masque les lignes vides dans des plages sp�cifiques de la feuille "6.35".
'             - C11:C39 : Plage de calcul (d�pend d'autres feuilles, g�r� par Worksheet_Calculate)
'             - C42:I52 : Plage de saisie (g�r� par Worksheet_Change)
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows66()
Attribute HideEmptyRows66.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange "6.35", "C11:C39"
    HideEmptyRowsInRange "6.35", "C42:I52", -1
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows68
' Objectif  : Masque les lignes vides dans une plage sp�cifique de la feuille "6.40-1".
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows68()
Attribute HideEmptyRows68.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange "6.40-1", "D4:D16"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows69
' Objectif  : Masque les lignes vides dans des plages sp�cifiques de la feuille "6.40-2".
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows69()
Attribute HideEmptyRows69.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange "6.40-2", "D20:D28"
    HideEmptyRowsInRange "6.40-2", "D32:D40"
    HideEmptyRowsInRange "6.40-2", "D44:D52"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows610
' Objectif  : Masque les lignes vides pour une liste de plages nomm�es sur la feuille "6.40-3".
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows610()
Attribute HideEmptyRows610.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim rngList As Variant
    rngList = Array("_6.40_3_B_2", "_6.40_3_B_3", "_6.40_3_B_4", "_6.40_3_B_5", _
        "_6.40_3_B_6", "_6.40_3_B_7", "_6.40_3_B_8", "_6.40_3_B_9", "_6.40_3_B_10", _
        "_6.40_3_C_2", "_6.40_3_C_3", "_6.40_3_C_4", "_6.40_3_C_5", _
        "_6.40_3_C_6", "_6.40_3_C_7", "_6.40_3_C_8", "_6.40_3_C_9", "_6.40_3_C_10", _
        "_6.40_3_D_2", "_6.40_3_D_3", "_6.40_3_D_4", "_6.40_3_D_5", _
        "_6.40_3_D_6", "_6.40_3_D_7", "_6.40_3_D_8", "_6.40_3_D_9", "_6.40_3_D_10")
    HideEmptyRowsForNamedRanges "6.40-3", rngList
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows611
' Objectif  : Masque les lignes vides pour une liste de plages nomm�es sur la feuille "6.40-4".
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows611()
Attribute HideEmptyRows611.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim rngList As Variant
    rngList = Array("_6.40_4_B_2", "_6.40_4_B_3", "_6.40_4_B_4", "_6.40_4_B_5", _
        "_6.40_4_B_6", "_6.40_4_B_7", "_6.40_4_B_8", "_6.40_4_B_9", "_6.40_4_B_10", _
        "_6.40_4_C_2", "_6.40_4_C_3", "_6.40_4_C_4", "_6.40_4_C_5", _
        "_6.40_4_C_6", "_6.40_4_C_7", "_6.40_4_C_8", "_6.40_4_C_9", "_6.40_4_C_10", _
        "_6.40_4_D_2", "_6.40_4_D_3", "_6.40_4_D_4", "_6.40_4_D_5", _
        "_6.40_4_D_6", "_6.40_4_D_7", "_6.40_4_D_8", "_6.40_4_D_9", "_6.40_4_D_10")
    HideEmptyRowsForNamedRanges "6.40-4", rngList
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows53
' Objectif  : Stub temporaire pour compilation.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows53()
Attribute HideEmptyRows53.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows57
' Objectif  : Stub temporaire pour compilation.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows57()
Attribute HideEmptyRows57.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows496
' Objectif  : Masque les lignes vides dans la plage B15:B34 de la feuille "RI".
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows496()
Attribute HideEmptyRows496.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange "RI", "B15:B34"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideEmptyRows70
' Objectif  : Masque les lignes vides dans la plage B18:D52 de la feuille Sheet70.
'---------------------------------------------------------------------------------------
Public Sub HideEmptyRows70()
Attribute HideEmptyRows70.VB_ProcData.VB_Invoke_Func = " \n14"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet70"), "B15:D24"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet70"), "B29:D38"
    HideEmptyRowsInRange GetSheetNameFromCodeName("Sheet70"), "B43:D52"
End Sub

' ===================================================================
' ===            HELPERS LOCALISES (STEP 11.4)                  ===
' ===================================================================

'---------------------------------------------------------------------------------------
' Fonction : TrySetWorksheetScrollArea (Step 11.4)
' Objectif : D�finit la ScrollArea d'une feuille de mani�re s�curis�e.
' Param�tres:
'   sheetName: Nom de la feuille
'   scrollArea: Zone de d�filement � d�finir
'   outError: Message d'erreur en cas d'�chec
' Retour: True si succ�s, False sinon
'---------------------------------------------------------------------------------------
Private Function TrySetWorksheetScrollArea(ByVal sheetName As String, ByVal scrollArea As String, ByRef outError As String) As Boolean
    On Error Resume Next
    outError = vbNullString

    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ThisWorkbook.sheets(sheetName)

    If Err.Number <> 0 Then
        outError = "Worksheet not found: " & Err.Description
        Err.Clear
        TrySetWorksheetScrollArea = False
        Exit Function
    End If

    If ws Is Nothing Then
        outError = "Worksheet is Nothing"
        TrySetWorksheetScrollArea = False
        Exit Function
    End If

    ws.scrollArea = scrollArea

    If Err.Number <> 0 Then
        outError = Err.Description
        Err.Clear
        TrySetWorksheetScrollArea = False
        Exit Function
    End If

    TrySetWorksheetScrollArea = True
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Fonction : TrySetColumnsHidden (Step 11.4)
' Objectif : Masque/affiche des colonnes de mani�re s�curis�e.
' Param�tres:
'   targetSheet: Feuille cible
'   columnRange: Plage de colonnes (ex: "A:C")
'   hidden: True pour masquer, False pour afficher
'   outError: Message d'erreur en cas d'�chec
' Retour: True si succ�s, False sinon
'---------------------------------------------------------------------------------------
Private Function TrySetColumnsHidden(ByVal targetSheet As Worksheet, ByVal columnRange As String, ByVal hidden As Boolean, ByRef outError As String) As Boolean
    On Error Resume Next
    outError = vbNullString

    targetSheet.Columns(columnRange).hidden = hidden

    If Err.Number <> 0 Then
        outError = Err.Description
        Err.Clear
        TrySetColumnsHidden = False
        Exit Function
    End If

    TrySetColumnsHidden = True
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddTrackingButtonToWorksheets
' Objectif  : Ajoute un bouton "Vue Suivi" sur toutes les feuilles de suivi définies.
'             Permet de basculer rapidement la vue sans passer par le ruban.
'---------------------------------------------------------------------------------------
Public Sub AddTrackingButtonToWorksheets()
    Dim ws As Worksheet
    Dim btn As Button
    Dim trackingSheets As Variant
    Dim i As Long
    
    Const BTN_NAME As String = "btnToggleViewLocal"
    Const MACRO_NAME As String = "'modSGQViews.ToggleViewMode'"
    
    ' Récupération sécurisée de la liste des feuilles
    On Error Resume Next
    trackingSheets = modConstants.TRACKING_SHEETS()
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de la récupération des feuilles de suivi : " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    For i = LBound(trackingSheets) To UBound(trackingSheets)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(trackingSheets(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' Supprimer l'ancien bouton s'il existe
            On Error Resume Next
            ws.Buttons(BTN_NAME).Delete
            On Error GoTo 0
            
            ' Ajouter le nouveau bouton (Position H1 par exemple, ou proche du titre)
            Dim targetCell As Range
            On Error Resume Next
            Set targetCell = ws.Range("H1")
            On Error GoTo 0
            
            If Not targetCell Is Nothing Then
                Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, 120, 25)
                With btn
                    .Name = BTN_NAME
                    .Caption = "Basculer Vue"
                    .OnAction = MACRO_NAME
                    .PrintObject = False ' Ne pas imprimer le bouton
                End With
            End If
        End If
    Next i
    
    MsgBox "Boutons 'Basculer Vue' ajoutés sur les feuilles de suivi.", vbInformation, "Configuration Interface"
End Sub


' ===============================================================================
' ===            UX & DASHBOARD IMPROVEMENTS (STEP 25)              ===
' ===============================================================================

'---------------------------------------------------------------------------------------
' Procedure : SetupDashboard
' Objectif  : Crée ou réinitialise la feuille 'Accueil' avec un design moderne.
'             Affiche les boutons de navigation principaux (SGQ, Suivi) et les paramètres.
'---------------------------------------------------------------------------------------
Public Sub SetupDashboard()
    Dim appScope As Object
    Set appScope = BeginAppStateScope("SetupDashboard")
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim shp As Shape
    Dim targetSheetName As String
    targetSheetName = "Accueil"

    ' 1. Créer ou récupérer la feuille Accueil
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo Handler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = targetSheetName
    Else
        If ws.Index <> 1 Then ws.Move Before:=ThisWorkbook.Sheets(1)
        ws.Activate
    End If

    ' 2. Nettoyage et Style
    With ws
        .Unprotect ' Si protégé
        .Cells.Clear
        .Cells.Interior.Color = RGB(245, 247, 250) ' Soft Gray
        .Cells.Font.Name = "Segoe UI"
        .DisplayPageBreaks = False
    End With
    
    ' Supprimer toutes les formes existantes
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' Masquer quadrillage et en-têtes
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100

    ' 3. Titre
    With ws.Range("B2")
        .Value = "CPA AI - CONSOLE"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80) ' Midnight Blue
    End With
    
    With ws.Range("B3")
        .Value = "Système de Gestion de la Qualité v" & GetCellValue("Entete", "A1")
        .Font.Size = 12
        .Font.Color = RGB(127, 140, 141) ' Concrete Gray
    End With

    ' 4. Boutons (Formes)
    
    ' --- Bouton 1 : SYSTÈME DE GESTION (SGQ) ---
    CreateDashboardButton ws, "btnSGQ", "SGQ System", _
        "💼" & vbCrLf & "Système de Gestion", _
        100, 100, 220, 120, _
        RGB(52, 73, 94), "'modSGQViews.GoToTDM'"

    ' --- Bouton 2 : SUIVI ---
    CreateDashboardButton ws, "btnTracking", "Tracking System", _
        "🖥️" & vbCrLf & "Activités de suivi", _
        340, 100, 220, 120, _
        RGB(22, 160, 133), "'modSGQViews.GoToTracking'"

    ' --- Bouton 3 : SETTINGS (Roue crantée) ---
    CreateDashboardButton ws, "btnSettings", "Settings", _
        "⚙️", _
        580, 100, 40, 40, _
        RGB(149, 165, 166), "modSGQAdministration.ToggleAdminMode"

    ' Protection
    ' ws.Protect UserInterfaceOnly:=True ' A voir si on active la protection tout de suite

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQInterface", "SetupDashboard", Err.Number, Err.Description
    Resume CleanExit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApplyVisualPolish
' Objectif  : Applique les standards visuels à TOUTES les feuilles (Zoom 100, Pas de grille).
'---------------------------------------------------------------------------------------
Public Sub ApplyVisualPolish()
    Dim appScope As Object
    Set appScope = BeginAppStateScope("ApplyVisualPolish")
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim currentWs As Worksheet
    Set currentWs = ActiveSheet
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = 100
            Application.Goto ws.Range("A1"), True
        End If
    Next ws
    
    If Not currentWs Is Nothing Then currentWs.Activate
    
    Application.ScreenUpdating = True

CleanExit:
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQInterface", "ApplyVisualPolish", Err.Number, Err.Description
    Resume CleanExit
End Sub

'---------------------------------------------------------------------------------------
' Helper : CreateDashboardButton
' Objectif : Crée un bouton stylisé (Shape) sur le dashboard.
'---------------------------------------------------------------------------------------
Private Sub CreateDashboardButton(ws As Worksheet, btnName As String, altText As String, _
                                  caption As String, left As Single, top As Single, _
                                  width As Single, height As Single, _
                                  backColor As Long, macroName As String)
    
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    
    With shp
        .Name = btnName
        .AlternativeText = altText
        .OnAction = macroName
        
        ' Style
        .Fill.ForeColor.RGB = backColor
        .Line.Visible = msoFalse
        .Shadow.Type = msoShadow25
        .ControlFormat.PrintObject = True
        
        ' Text
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = caption
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 14
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub
��"(9e377b08b88e30057baf56c074b3d12f1d2237232>file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQInterface.bas:file:///c:/VBA/SGQ%201.65