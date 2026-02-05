Attribute VB_Name = "TestGraph"
Option Explicit

' ----------------------------------------------------
' Script de Test de Connexion Microsoft Graph (VBA)
' Authentification "App Only" (Client Credentials)
' ----------------------------------------------------

Sub TestConnection()
    ' --- CONFIGURATION ---
    Dim tenantId As String: tenantId = "fe13311a-8cd8-4985-ad94-b2537442986e"
    Dim clientId As String: clientId = "334b3889-cd27-40cf-b101-6b35d25f57cb"
    
    ' [IMPORTANT] Collez ici le Secret que vous avez copié plus tôt
    Dim clientSecret As String
    clientSecret = "J5k8Q~Jo5Uax_nt_OygWEyfo4h_oHsNjBQ-h9b48" 
    
    If clientSecret = "" Or clientSecret = "YOUR_SECRET_HERE" Then
        MsgBox "Veuillez renseigner la variable clientSecret dans le code.", vbExclamation
        Exit Sub
    End If
    
    ' 1. Récupération du Token
    Debug.Print "--- DÉBUT DU TEST ---"
    Debug.Print "Tentative d'obtention du token..."
    
    Dim token As String
    token = GetGraphToken(tenantId, clientId, clientSecret)
    
    If token = "" Then
        MsgBox "ECHEC : Impossible d'obtenir le token. Vérifiez la fenêtre Exécution (Ctrl+G).", vbCritical
        Exit Sub
    End If
    
    Debug.Print "SUCCÈS : Token obtenu !"
    Debug.Print "Debut du token: " & Left(token, 20) & "..."
    
    ' 2. Test d'appel API (Vérification du téléchargement direct)
    ' URL exacte qui échoue dans modSecureUpdate (Status 0)
    Debug.Print "Test URL: .../manifest.json:/content"
    
    Dim response As String
    response = CallGraphApi("https://graph.microsoft.com/v1.0/sites/root/drive/root:/General/SGQ_Updates/Production/manifest.json:/content", "GET", token)
    
    Debug.Print "Réponse API brute :"
    Debug.Print response
    
    If InStr(response, "error") > 0 Or response = "" Then
        MsgBox "Token valide, mais l'appel API a échoué. Voir Exécution.", vbExclamation
    Else
        MsgBox "SUCCÈS TOTAL ! Connexion Graph opérationnelle.", vbInformation
    End If
    Debug.Print "--- FIN DU TEST ---"
End Sub

' --- FONCTIONS UTILITAIRES ---

Function GetGraphToken(tenantId, clientId, clientSecret) As String
    Dim objHTTP As Object
    ' Utilisation de MSXML2 pour compatibilité maximale
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    
    Dim url As String
    url = "https://login.microsoftonline.com/" & tenantId & "/oauth2/v2.0/token"
    
    Dim body As String
    ' Note: scope=.default est obligatoire pour le flux Client Credentials
    body = "client_id=" & clientId & "&scope=https://graph.microsoft.com/.default&client_secret=" & clientSecret & "&grant_type=client_credentials"
    
    On Error Resume Next
    objHTTP.Open "POST", url, False
    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHTTP.send body
    
    If Err.Number <> 0 Then
        Debug.Print "Erreur réseau : " & Err.Description
        GetGraphToken = ""
        Exit Function
    End If
    On Error GoTo 0
    
    If objHTTP.Status = 200 Then
        ' Parsing manuel simple pour éviter d'importer JsonConverter pour ce test
        Dim json As String: json = objHTTP.responseText
        Dim startPos As Long: startPos = InStr(json, """access_token"":""") + 16
        If startPos > 16 Then
            Dim endPos As Long: endPos = InStr(startPos, json, """")
            GetGraphToken = Mid(json, startPos, endPos - startPos)
        Else
            GetGraphToken = ""
        End If
    Else
        Debug.Print "Erreur HTTP " & objHTTP.Status & ": " & objHTTP.responseText
        GetGraphToken = ""
    End If
End Function

Function CallGraphApi(url As String, method As String, token As String) As String
    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    
    On Error Resume Next
    objHTTP.Open method, url, False
    objHTTP.setRequestHeader "Authorization", "Bearer " & token
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send
    
    If Err.Number <> 0 Then
        CallGraphApi = "Erreur réseau: " & Err.Description
    Else
        CallGraphApi = objHTTP.responseText
    End If
End Function
