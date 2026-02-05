°
' -------------------------------------------------------------------------
' Secure File Operations (Added Backlog 2)
' -------------------------------------------------------------------------

Public Function ReadFileContent(ByVal filePath As String) As String
    ' Reads a file. Handles text files safely.
    On Error GoTo Handler
    Dim content As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        ReadFileContent = ""
        Exit Function
    End If
    
    ' Read as System Default (TristateUseDefault = -2) -> Handles CP1252 usually
    Set ts = fso.OpenTextFile(filePath, 1, False, -2)
    If Not ts.AtEndOfStream Then
        content = ts.ReadAll
    End If
    ts.Close
    
    ReadFileContent = content
    Exit Function
    
Handler:
    ' LogError is in modSGQUtilitaires
    modSGQUtilitaires.LogError "ReadFileContent failed for: " & filePath
    ReadFileContent = ""
End Function

Public Sub WriteFileContent(ByVal filePath As String, ByVal content As String)
    ' Writes content to file with automatic .bak backup.
    On Error GoTo Handler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Backup if exists
    If fso.FileExists(filePath) Then
        Dim bakPath As String
        bakPath = filePath & ".bak"
        On Error Resume Next ' Ignore error if we can't delete old backup
        If fso.FileExists(bakPath) Then fso.DeleteFile bakPath, True
        On Error GoTo Handler
        fso.CopyFile filePath, bakPath, True
    End If
    
    ' 2. Write content
    Dim ts As Object
    ' CreateTextFile(Filename, Overwrite, Unicode) -> Unicode=False implies ASCII/ANSI
    Set ts = fso.CreateTextFile(filePath, True, False) 
    ts.Write content
    ts.Close
    
    Exit Sub
Handler:
    modSGQUtilitaires.LogError "WriteFileContent failed for: " & filePath
    Err.Raise Err.Number, "WriteFileContent", "Error writing file: " & Err.Description
End Sub
°*cascade08"(b3eea256776ced57c23cbf87c776d2461e607ebe2@file:///C:/Users/AbelBoudreau/Workspace_CPA_AI/temp_fs_funcs.bas:.file:///C:/Users/AbelBoudreau/Workspace_CPA_AI