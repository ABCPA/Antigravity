�\Attribute VB_Name = "modSGQViews"
Option Explicit

' --------------------------------------------------------------------------------------------------------------------
' Module    : modSGQViews
' Auteur    : Abel Boudreau / Antigravity
' Date      : 2026-01-27
' Objectif  : Gère les modes d'affichage du classeur (Mode Complet vs Mode Suivi).
' Description: Permet de basculer entre une vue "Système" (audit complet) et une vue "Suivi" (annuel)
'              en masquant/affichant les feuilles pertinentes.
' --------------------------------------------------------------------------------------------------------------------

Private Const VIEW_MODE_NAME As String = "SGQ_CurrentViewMode"
Private Const MODE_SYSTEM As String = "SYSTEM"
Private Const MODE_TRACKING As String = "TRACKING"

' --------------------------------------------------------------------------------------------------------------------
' Procédure : ToggleViewMode
' Objectif  : Bascule entre le mode Système et le mode Suivi.
' --------------------------------------------------------------------------------------------------------------------
Public Sub ToggleViewMode()
    Dim currentMode As String
    currentMode = GetCurrentViewMode()
    
    Dim appScope As Object
    Set appScope = BeginAppStateScope("ToggleViewMode")
    On Error GoTo Handler
    
    Application.ScreenUpdating = False ' EVITER FLICKERING
    
    If currentMode = MODE_TRACKING Then
        SetViewMode MODE_SYSTEM
        MsgBox "Mode Systeme active. Toutes les feuilles sont visibles.", vbInformation, "Vue SGQ"
    Else
        SetViewMode MODE_TRACKING
        MsgBox "Mode Suivi active. Seules les feuilles de suivi sont visibles.", vbInformation, "Vue SGQ"
    End If
    
    ' Activer la première feuille visible pour éviter les erreurs de sélection
    ActivateFirstVisibleSheet ThisWorkbook
    
CleanExit:
    Application.ScreenUpdating = True
    If Not appScope Is Nothing Then Set appScope = Nothing
    Exit Sub
Handler:
    LogError "modSGQViews", "ToggleViewMode", Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Sub ActivateFirstVisibleSheet(wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            Exit Sub
        End If
    Next ws
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------------------------------------------------
' Procédure : SetViewMode
' Objectif  : Applique un mode d'affichage spécifique.
' --------------------------------------------------------------------------------------------------------------------
Public Sub SetViewMode(ByVal modeName As String)
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Enregistre l'état
    modExcelUtils.SetNamedValue VIEW_MODE_NAME, modeName, wb
    
    Select Case modeName
        Case MODE_TRACKING
            ApplyTrackingMode wb
        Case MODE_SYSTEM
            ApplySystemMode wb
        Case Else
            ' Par défaut, tout afficher
            ApplySystemMode wb
    End Select
End Sub

' --------------------------------------------------------------------------------------------------------------------
' Fonction  : GetCurrentViewMode
' Objectif  : Récupère le mode actuel depuis le nom défini.
' --------------------------------------------------------------------------------------------------------------------
Public Function GetCurrentViewMode() As String
    Dim val As String
    If modExcelUtils.TryGetNamedValue(VIEW_MODE_NAME, val, ThisWorkbook) Then
        GetCurrentViewMode = val
    Else
        GetCurrentViewMode = MODE_SYSTEM ' Defaut
    End If
End Function

' --------------------------------------------------------------------------------------------------------------------
' Procédure : ApplyTrackingMode
' Objectif  : Masque toutes les feuilles SAUF celles du suivi.
' --------------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------------
' Procédure : ApplyTrackingMode
' Objectif  : Masque toutes les feuilles SAUF celles du suivi.
'             Si Admin = False, masque STRICTEMENT les feuilles techniques.
' --------------------------------------------------------------------------------------------------------------------
Private Sub ApplyTrackingMode(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim trackingSheets As Variant
    Dim technicalSheets As Variant
    Dim isAdmin As Boolean
    
    trackingSheets = modConstants.TRACKING_SHEETS()
    technicalSheets = modConstants.TECHNICAL_SHEETS()
    isAdmin = modSGQAdministration.isAdminModeActive()
    
    ' 1. Identifier les feuilles à RENDRE VISIBLES (Liste Suivi - Techniques si pas admin)
    ' On fait d'abord une boucle pour afficher celles requises
    Dim i As Long
    Dim shouldBeVisible As Boolean
    Dim firstVisible As Boolean
    firstVisible = False
    
    For Each ws In wb.Worksheets
        shouldBeVisible = False
        
        ' Est-ce une feuille de suivi ?
        If IsInArray(ws.Name, trackingSheets) Then
            shouldBeVisible = True
        End If
        
        ' Si c'est une feuille technique et qu'on n'est PAS admin -> Masquer
        If IsInArray(ws.Name, technicalSheets) And Not isAdmin Then
            shouldBeVisible = False
        End If
        
        If shouldBeVisible Then
            ws.Visible = xlSheetVisible
            firstVisible = True
        End If
    Next ws
    
    If Not firstVisible Then
        ' Fallback : TDM-Suivi ou première feuille de suivi
        On Error Resume Next
        wb.Sheets("TDM-Suivi").Visible = xlSheetVisible
        On Error GoTo 0
    End If
    
    ' 2. Masquer tout le reste (Strictement)
    For Each ws In wb.Worksheets
        shouldBeVisible = False
        If IsInArray(ws.Name, trackingSheets) Then shouldBeVisible = True
        If IsInArray(ws.Name, technicalSheets) And Not isAdmin Then shouldBeVisible = False
        
        ' Si Admin est actif, on laisse les feuilles techniques visibles SI elles ont été affichées manuellement
        ' Mais ici on applique un mode, donc on reset tout ce qui n'est pas "Suivi"
        
        If Not shouldBeVisible Then
            ' Si Admin, on peut laisser visible si c'est technique ?
            ' La règle demandée : "Les faire disparaître en mode suivi" sauf si admin.
            ' Si Admin, on ne force pas le masquage des techniques ?
            ' "Les onglets Calcul... masquées, SAUF lorsque le mode admin est activé"
            
            If isAdmin And IsInArray(ws.Name, technicalSheets) Then
                ' On ne change pas l'état actuel (si l'utilisateur l'a affiché) ou on l'affiche ?
                ' Pour l'instant, masquons par défaut en mode suivi, l'admin pourra les afficher via "Activer Admin"
                ' ou on les affiche si on bascule ?
                ' La demande est "Doivent toujours être masquées sauf en admin".
                ' En mode Suivi, on veut voir le Suivi.
                ws.Visible = xlSheetVeryHidden
            Else
                ws.Visible = xlSheetVeryHidden
            End If
        End If
    Next ws
End Sub

' --------------------------------------------------------------------------------------------------------------------
' Procédure : ApplySystemMode
' Objectif  : Affiche les feuilles SYSTEM (Excluant Suivi, Excluant Technique sauf Admin).
' --------------------------------------------------------------------------------------------------------------------
Private Sub ApplySystemMode(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim systemSheets As Variant
    Dim technicalSheets As Variant
    Dim trackingSheets As Variant
    Dim isAdmin As Boolean
    
    systemSheets = modConstants.SYSTEM_SHEETS()
    technicalSheets = modConstants.TECHNICAL_SHEETS()
    trackingSheets = modConstants.TRACKING_SHEETS()
    isAdmin = modSGQAdministration.isAdminModeActive()
    
    For Each ws In wb.Worksheets
        ' Logique d'affichage
        ' On affiche si c'est dans SYSTEM
        ' On masque si c'est dans TRACKING (et pas dans SYSTEM, mais disjoint normalement)
        ' On masque si c'est TECHNIQUE et pas Admin
        
        Dim showIt As Boolean
        showIt = False
        
        If IsInArray(ws.Name, systemSheets) Then showIt = True
        
        ' Exclusion stricte Suivi (si jamais il y a chevauchement, Suivi gagne en masquage ?)
        ' "Les onglets du Mode Suivi ne doivent pas être visible en Mode SGQ"
        If IsInArray(ws.Name, trackingSheets) Then showIt = False
        
        ' Gestion Technique
        If IsInArray(ws.Name, technicalSheets) Then
            If isAdmin Then
                ' En mode Admin, on peut voir les techniques
                ' showIt reste True si c'était dans System
                ' Si c'était pas dans System (ex: EnteteSuivi), on l'affiche quand même si Admin ? Non, contexte.
                ' Si la feuille est technique et appartient au contexte courant, on l'affiche si admin.
                If showIt Then showIt = True
            Else
                showIt = False
            End If
        End If
        
        If showIt Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
End Sub

' --------------------------------------------------------------------------------------------------------------------
' Utilitaires
' --------------------------------------------------------------------------------------------------------------------
Private Function IsInArray(ByVal value As String, ByVal arr As Variant) As Boolean
    Dim element As Variant
    On Error Resume Next
    For Each element In arr
        If UCase(CStr(element)) = UCase(value) Then
            IsInArray = True
            Exit Function
        End If
    Next element
    On Error GoTo 0
    IsInArray = False
End Function

' --------------------------------------------------------------------------------------------------------------------
' Procédure : GoToTDM
' Objectif  : Active le mode Système et navigue vers la feuille TDM.
'             Utilisé par le bouton du Dashboard.
' --------------------------------------------------------------------------------------------------------------------
Public Sub GoToTDM()
    SetViewMode MODE_SYSTEM
    On Error Resume Next
    ThisWorkbook.Sheets("TDM").Visible = xlSheetVisible
    ThisWorkbook.Sheets("TDM").Activate
    If Err.Number <> 0 Then
        MsgBox "La feuille 'TDM' est introuvable.", vbExclamation, "Navigation"
    End If
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------------------------------------------------
' Procédure : GoToTracking
' Objectif  : Active le mode Suivi et navigue vers la feuille TDM-Suivi.
'             Utilisé par le bouton du Dashboard.
' --------------------------------------------------------------------------------------------------------------------
Public Sub GoToTracking()
    SetViewMode MODE_TRACKING
    On Error Resume Next
    ThisWorkbook.Sheets("TDM-Suivi").Visible = xlSheetVisible
    ThisWorkbook.Sheets("TDM-Suivi").Activate
    If Err.Number <> 0 Then
        MsgBox "La feuille 'TDM-Suivi' est introuvable.", vbExclamation, "Navigation"
    End If
    On Error GoTo 0
End Sub
�\"(9e377b08b88e30057baf56c074b3d12f1d2237232:file:///c:/VBA/SGQ%201.65/vba-files/Module/modSGQViews.bas:file:///c:/VBA/SGQ%201.65