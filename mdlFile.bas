Attribute VB_Name = "mdlFile"
Option Explicit

'***File functions***

Public Function NormalizeDir(ByVal sDir As String) As String
    If Not Right$(sDir, 1) = "\" Then sDir = sDir & "\"
    NormalizeDir = sDir
End Function

Public Function GetDir(ByVal sPath As String) As String
    GetDir = NormalizeDir(Left$(sPath, Len(sPath) - InStr(1, StrReverse(sPath), "\") + 1))
End Function

Public Function GetFile(ByVal sPath As String) As String
    If InStr(sPath, "\") = 0 Then
        GetFile = sPath
    Else
        GetFile = Right$(sPath, InStr(1, StrReverse(sPath), "\") - 1)
    End If
End Function

Public Function GetFileWOExt(ByVal sPath As String) As String
    sPath = GetFile(sPath)
    If InStr(sPath, ".") = 0 Then
        GetFileWOExt = sPath
    Else
        GetFileWOExt = Left$(sPath, Len(sPath) - InStr(1, StrReverse(sPath), "."))
    End If
End Function
