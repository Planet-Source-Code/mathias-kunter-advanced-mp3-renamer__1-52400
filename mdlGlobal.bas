Attribute VB_Name = "mdlGlobal"
Option Explicit

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Enum MP3SourceEnum
    SOURCE_IDV1
    SOURCE_IDV2
    SOURCE_FILENAME
    SOURCE_USERENTRY
End Enum

Public Type MP3File
    SourceFile As String
    SourceType As MP3SourceEnum
    
    FileInterpretItems() As String
    FileInterpretItemCnt As Long
    FileInterpretArtist As Long
    FileInterpretTitle As Long
      
    HasIDv1 As Boolean
    HasIDv2 As Boolean
    IDv1 As ID3Tag
    IDv2 As ID3Tag
    FileTag As ID3Tag
    UserTag As ID3Tag
End Type

'APIs for the folder selection
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'APIs used to find files.
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'Renaming-API
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long


Public RdFiles() As MP3File, nFiles As Long


Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Public Function FindFilesAPI(ByVal Path As String, ByVal SearchStr As String, ByRef outFiles() As String, ByVal SubDirs As Boolean, Optional ByRef nFiles As Long = 0) As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    
    Path = NormalizeDir(Path)
    
    'Walk through this directory and get matching files.
    hSearch = FindFirstFile(Path & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                'Valid item.
                If SubDirs And (GetFileAttributes(Path & CurItem) And FILE_ATTRIBUTE_DIRECTORY) Then
                    'Item is a sub-directory, read it recursivly.
                    FindFilesAPI Path & CurItem, SearchStr, outFiles(), True, nFiles
                ElseIf InStr(1, Path & CurItem, SearchStr, vbTextCompare) Then
                    'Item is a file which we're searching for.
                    ReDim Preserve outFiles(nFiles)
                    outFiles(nFiles) = Path & CurItem
                    nFiles = nFiles + 1
                End If
            End If
            'Get next item
            Result = FindNextFile(hSearch, WFD)
        Loop
        FindClose hSearch
    End If
    'Return the number of files in this directory (well, is also stored in the ByRef parameter nFiles).
    FindFilesAPI = nFiles
End Function

Public Function CleanStr(ByVal strData As String, ByVal UpperCase As Boolean, ByVal LowerCase As Boolean, ByVal CutLeadingNumber As Boolean) As String
    Dim i As Long
    Dim SplitField() As String, NewStr As String
    
    'Replacing...
    strData = ReplaceStr(strData)
    'Trim the string.
    CleanStr = Trim$(strData)
    If CleanStr = "" Then Exit Function
    'Remove multiple spaces.
    Do While Not InStr(1, CleanStr, "  ", vbTextCompare) = 0
        CleanStr = Replace$(CleanStr, "  ", " ", , , vbTextCompare)
    Loop
    'Upper case and / or lower case the string correctly.
    SplitField = Split(CleanStr, " ", , vbTextCompare)
    CleanStr = ""
    For i = 0 To UBound(SplitField)
        If Not i = 0 Or Not CutLeadingNumber Or Not IsNumeric(SplitField(i)) Then
            If UpperCase Then
                NewStr = UCase$(Left$(SplitField(i), 1))
            Else
                NewStr = Left$(SplitField(i), 1)
            End If
            If LowerCase Then
                NewStr = NewStr & LCase$(Right$(SplitField(i), Len(SplitField(i)) - 1))
            Else
                NewStr = NewStr & Right$(SplitField(i), Len(SplitField(i)) - 1)
            End If
            CleanStr = CleanStr & NewStr & " "
        End If
    Next i
    CleanStr = Trim$(CleanStr)
End Function

Private Function ReplaceStr(ByVal strData As String) As String
    'Replace invalid sings.
    strData = Replace$(strData, "_", " ", , , vbTextCompare)
    strData = Replace$(strData, "´", "'", , , vbTextCompare)
    strData = Replace$(strData, "`", "'", , , vbTextCompare)
    strData = Replace$(strData, "{", "(", , , vbTextCompare)
    strData = Replace$(strData, "[", "(", , , vbTextCompare)
    strData = Replace$(strData, "]", ")", , , vbTextCompare)
    strData = Replace$(strData, "}", ")", , , vbTextCompare)
    'Cut out invalid signs.
    strData = Replace$(strData, "/", "", , , vbTextCompare)
    strData = Replace$(strData, "\", "", , , vbTextCompare)
    strData = Replace$(strData, ":", "", , , vbTextCompare)
    strData = Replace$(strData, "*", "", , , vbTextCompare)
    strData = Replace$(strData, "?", "", , , vbTextCompare)
    strData = Replace$(strData, """", "", , , vbTextCompare)
    strData = Replace$(strData, "<", "", , , vbTextCompare)
    strData = Replace$(strData, ">", "", , , vbTextCompare)
    strData = Replace$(strData, "|", "", , , vbTextCompare)
    ReplaceStr = strData
End Function

Public Function SplitInterpreteItems(ByVal strData As String, ByVal UpperCase As Boolean, ByVal LowerCase As Boolean, ByRef outField() As String) As Long
    Dim i As Long
    Dim WorkStr As String, StrField() As String
    Dim outCnt As Long
    
    'Replace "___" with "-".
    WorkStr = Replace$(strData, "___", " - ", , , vbTextCompare)
    
    'Check the parts between two "-". Remove a part if it's numeric or an album name.
    StrField = Split(WorkStr, "-", , vbTextCompare)
    WorkStr = ""
    For i = 0 To UBound(StrField)
        'Adjust every string part of its own.
        StrField(i) = Trim$(StrField(i))
        If i = 0 Then
            StrField(i) = CleanStr(StrField(i), UpperCase, LowerCase, False)
        Else
            StrField(i) = CleanStr(StrField(i), UpperCase, LowerCase, True)
        End If
        If Not StrField(i) = "" Then
            If Not IsNumeric(StrField(i)) Then
                'This is a valid entry.
                ReDim Preserve outField(outCnt)
                outField(outCnt) = StrField(i)
                outCnt = outCnt + 1
            End If
        End If
    Next i
    SplitInterpreteItems = outCnt
End Function

Public Function CleanInterpreteItems(ByVal strData As String) As String
    Dim i As Long
    Dim WorkStr As String, StrField() As String
    
    'Replace "___" with "-".
    WorkStr = Replace$(strData, "___", " - ", , , vbTextCompare)
    
    'Check the parts between two "-". Remove a part if it's numeric or an album name.
    StrField = Split(WorkStr, "-", , vbTextCompare)
    WorkStr = ""
    For i = 0 To UBound(StrField)
        'Adjust every string part of its own.
        StrField(i) = Trim$(StrField(i))
        StrField(i) = CleanStr(StrField(i), False, False, False)
        If Not StrField(i) = "" Then
            If Not IsNumeric(StrField(i)) Then
                CleanInterpreteItems = CleanInterpreteItems & StrField(i) & "-"
            End If
        End If
    Next i
    'Remove the final "-".
    If Not Len(CleanInterpreteItems) = 0 Then CleanInterpreteItems = Left$(CleanInterpreteItems, Len(CleanInterpreteItems) - 1)
End Function
