øAttribute VB_Name = "modTest_Utilitaires"
Option Explicit

Public Sub RunTests()
    ' Called by modUnitTestEngine
    Test_SanitizeFileName
    Test_IsVbideTrusted
End Sub

Private Sub Test_SanitizeFileName()
    Dim inputStr As String
    Dim expected As String
    Dim actual As String
    
    ' Case 1: Simple
    inputStr = "File:Name?.xlsx"
    expected = "File_Name_.xlsx"
    actual = modSGQFileSystem.SanitizeFileNamePart(inputStr)
    modUnitTestEngine.AssertEqual expected, actual, "SanitizeFileName: Special chars"
    
    ' Case 2: Clean
    inputStr = "CleanName"
    expected = "CleanName"
    actual = modSGQFileSystem.SanitizeFileNamePart(inputStr)
    modUnitTestEngine.AssertEqual expected, actual, "SanitizeFileName: No change"
End Sub

Private Sub Test_IsVbideTrusted()
    ' Just ensure it runs without error and returns a boolean
    Dim res As Boolean
    On Error Resume Next
    res = modSGQUtilitaires.IsVbideTrusted()
    modUnitTestEngine.AssertTrue (Err.Number = 0), "IsVbideTrusted: Runs safely"
    ' Result depends on environment, so we just asserts it didn't crash
End Sub
ø"(9e377b08b88e30057baf56c074b3d12f1d2237232Bfile:///c:/VBA/SGQ%201.65/vba-files/Module/modTest_Utilitaires.bas:file:///c:/VBA/SGQ%201.65