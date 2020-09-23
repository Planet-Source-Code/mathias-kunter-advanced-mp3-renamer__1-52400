Attribute VB_Name = "mdlID3"
Option Explicit

Public Type ID3Tag
    Artist As String
    Title As String
    Album As String
End Type


'ID3v1 Types
Private Type ID3v1Tag
    Identifier(2) As Byte
    Title(29) As Byte
    Artist(29) As Byte
    Album(29) As Byte
    SongYear(3) As Byte
    Comment(29) As Byte
    Genre As Byte
End Type

'ID3v2 Types
Private Type ID3v2Header
    Identifier(2) As Byte
    Version(1) As Byte
    Flags As Byte
    Size(3) As Byte
End Type

Private Type ID3v2ExtendedHeader
    Size(3) As Byte
End Type

Private Type ID3v2FrameHeader
    FrameID(3) As Byte
    Size(3) As Byte
    Flags(1) As Byte
End Type


Public Function ReadID3v1(ByVal strFile As String, ByRef outTag As ID3Tag) As Boolean
    Dim FileNo As Integer, fp As Long
    Dim RdTag As ID3v1Tag
    
    On Local Error GoTo Failed
    
    FileNo = FreeFile
    Open strFile For Binary As #FileNo
        fp = LOF(FileNo) - 127
        If fp > 0 Then
            Get #FileNo, fp, RdTag
            If GetStringValue(RdTag.Identifier, 3, 0) = "TAG" Then
                'An ID3v1 tag is present.
                outTag.Artist = Trim$(GetStringValue(RdTag.Artist, 30, 0))
                outTag.Title = Trim$(GetStringValue(RdTag.Title, 30, 0))
                outTag.Album = Trim$(GetStringValue(RdTag.Album, 30, 0))
                ReadID3v1 = True
            End If
        End If
    Close #FileNo
    Exit Function
    
Failed:
    ReadID3v1 = False
End Function

Public Function WriteID3v1(ByVal strFile As String, ByRef outTag As ID3Tag) As Boolean
    Dim FileNo As Integer, fp As Long
    Dim RdTag As ID3v1Tag, WrTag As ID3v1Tag, LocalTag As ID3Tag
    
    On Local Error GoTo Failed
    
    If outTag.Artist = "" And outTag.Title = "" And outTag.Album = "" Then Exit Function
    
    'Writes the ID3v1 tag of an mp3 file.
    FileNo = FreeFile
    Open strFile For Binary As #FileNo
        fp = LOF(FileNo) - 127
        If fp > 0 Then
            Get #FileNo, fp, RdTag
            If GetStringValue(RdTag.Identifier, 3, 0) = "TAG" Then
                fp = LOF(FileNo) - 127
            Else
                'An ID3v1 tag is not present.
                fp = LOF(FileNo) + 1
            End If
        Else
            'Not really needed... which file is smaller than 128 bytes?
            'But it's better coding...
            fp = LOF(FileNo) + 1
        End If
        LocalTag.Artist = outTag.Artist
        LocalTag.Title = outTag.Title
        LocalTag.Album = outTag.Album
        If Len(LocalTag.Artist) > 30 Then LocalTag.Artist = Left$(LocalTag.Artist, 30)
        If Len(LocalTag.Title) > 30 Then LocalTag.Title = Left$(LocalTag.Title, 30)
        If Len(LocalTag.Album) > 30 Then LocalTag.Album = Left$(LocalTag.Album, 30)
        SetStringValue WrTag.Identifier, "TAG", 3
        SetStringValue WrTag.Artist, LocalTag.Artist, Len(LocalTag.Artist)
        SetStringValue WrTag.Title, LocalTag.Title, Len(LocalTag.Title)
        SetStringValue WrTag.Album, LocalTag.Album, Len(LocalTag.Album)
        WrTag.Genre = 255
        Put #FileNo, fp, WrTag
    Close #FileNo
    WriteID3v1 = True
    Exit Function
    
Failed:
    WriteID3v1 = False
End Function

Public Function ReadID3v2(ByVal strFile As String, ByRef outTag As ID3Tag) As Boolean
    Dim i As Integer, FileNo As Integer, fp As Long
    Dim RdHeader As ID3v2Header, RdExtHeader As ID3v2ExtendedHeader, RdFrameHeader As ID3v2FrameHeader
    Dim FrameID As String, FrameSize As Long, TextEncoding As Byte, RdData() As Byte, RdString As String
    Dim bGotArtist As Boolean, bGotTitle As Boolean
    
    On Local Error GoTo Failed
    
    'Reads the ID3v2 tag of an mp3 file, if there is one.
    FileNo = FreeFile
    fp = 1
    Open strFile For Binary As #FileNo
        'Read the header.
        Get #FileNo, fp, RdHeader
        
        If GetStringValue(RdHeader.Identifier, 3, 0) = "ID3" Then
            fp = Loc(FileNo) + 1
            
            'An ID3v2 tag is present.
            If GetBit(6, RdHeader.Flags) Then
                'There is an extended header present. Just read its size to jump over it.
                Get #FileNo, , RdExtHeader
                fp = fp + GetLongValue(RdExtHeader.Size)
            End If
            
            Do
                Get #FileNo, fp, RdFrameHeader
                FrameID = GetStringValue(RdFrameHeader.FrameID, 4, 0)
                FrameSize = GetLongValue(RdFrameHeader.Size)
                If Not FrameSize < 2 Then
                    If FrameID = "TPE1" Or FrameID = "TIT2" Or FrameID = "TALB" Then
                        Get #FileNo, , TextEncoding
                        ReDim RdData(FrameSize - 2)
                        Get #FileNo, , RdData
                        RdString = GetStringValue(RdData, UBound(RdData) + 1, TextEncoding)
                        If FrameID = "TPE1" Then
                            'Artist frame.
                            outTag.Artist = RdString
                            bGotArtist = True
                        ElseIf FrameID = "TIT2" Then
                            'Title frame.
                            outTag.Title = RdString
                            bGotTitle = True
                        Else
                            'Album frame.
                            outTag.Album = RdString
                        End If
                    End If
                End If
                'Seek to the next frame. The value + 10 is the frame header itself.
                fp = fp + 10 + FrameSize
            Loop While Not FrameSize = 0 And Not fp > 10 + GetLongValue(RdHeader.Size)
            If bGotArtist And bGotTitle Then ReadID3v2 = True
        End If
    Close #FileNo
    Exit Function
    
Failed:
    ReadID3v2 = False
End Function

Public Function WriteID3v2(ByVal strFile As String, ByRef outTag As ID3Tag) As Boolean
    Dim i As Integer, FileNo As Integer, fp As Long
    Dim AudioData() As Byte, AudioSize As Long, TagSize As Long
    Dim Header As ID3v2Header, WrHeader As ID3v2Header
    
    On Local Error GoTo Failed
    
    TagSize = Len(outTag.Artist) + Len(outTag.Title) + Len(outTag.Album)
    If Not Len(outTag.Artist) = 0 Then TagSize = TagSize + 11
    If Not Len(outTag.Title) = 0 Then TagSize = TagSize + 11
    If Not Len(outTag.Album) = 0 Then TagSize = TagSize + 11
    
    'Writes the ID3v2 tag of an mp3 file.
    FileNo = FreeFile
    fp = 1
    Open strFile For Binary As #FileNo
        AudioSize = LOF(FileNo)
        'Check for an existing header.
        Get #FileNo, fp, Header
        If GetStringValue(Header.Identifier, 3, 0) = "ID3" Then
            AudioSize = AudioSize - GetLongValue(Header.Size)
        End If
        'Save the existing audio data.
        ReDim AudioData(AudioSize - 1)
        Get #FileNo, LOF(FileNo) - AudioSize + 1, AudioData
    Close #FileNo
    Kill strFile
    Open strFile For Binary As #FileNo
        'Create the ID3 tag.
        '1) Create the header.
        SetStringValue WrHeader.Identifier, "ID3", 3
        WrHeader.Version(0) = 3
        SetLongValue WrHeader.Size, TagSize
        Put #FileNo, , WrHeader
        '2) Create the frames.
        WriteFrame FileNo, "TPE1", outTag.Artist
        WriteFrame FileNo, "TIT2", outTag.Title
        WriteFrame FileNo, "TALB", outTag.Album
        '3) Append the audio data.
        Put #FileNo, , AudioData
    Close #FileNo
    
    WriteID3v2 = True
    Exit Function
    
Failed:
    WriteID3v2 = False
End Function

Private Sub WriteFrame(ByVal FileNo As Integer, ByVal strFrameHeader As String, ByVal strFrameData As String)
    Dim FrameHeader As ID3v2FrameHeader, EncData As Byte, FrameData() As Byte
    
    If Not Len(strFrameData) = 0 Then
        SetStringValue FrameHeader.FrameID, strFrameHeader, 4
        SetLongValue FrameHeader.Size, Len(strFrameData) + 1
        Put #FileNo, , FrameHeader
        ReDim FrameData(Len(strFrameData) - 1)
        SetStringValue FrameData, strFrameData, Len(strFrameData)
        Put #FileNo, , EncData
        Put #FileNo, , FrameData
    End If
End Sub

'Synchsafe integers are integers that keep its highest bit (bit 7) zeroed, making seven bits
'out of eight available. Thus a 32 bit synchsafe integer can store 28 bits of information.
Private Function GetLongValue(ByRef SyncsafeInt() As Byte) As Long
    Dim i As Integer, j As Integer, BitNr As Integer
    
    For i = 3 To 0 Step -1
        'Loop through the 4 bytes.
        For j = 0 To 6
            'Loop through the 7 significant bits per byte.
            If GetBit(j, SyncsafeInt(i)) Then
                GetLongValue = GetLongValue + 2 ^ BitNr
            End If
            BitNr = BitNr + 1
        Next j
    Next i
End Function

Private Sub SetLongValue(ByRef SyncsafeInt() As Byte, ByVal Value As Long)
    Dim i As Integer, ByteNr As Integer, BitNr As Integer
    
    ByteNr = 3
    For i = 0 To 27
        'Loop through the 28 bits of an synchsafe integer.
        If Value And 2 ^ i Then
            'This bit is set.
            SetBit BitNr, SyncsafeInt(ByteNr), True
        End If
        BitNr = BitNr + 1
        If BitNr Mod 7 = 0 Then
            'The next byte begins.
            ByteNr = ByteNr - 1
            BitNr = 0
        End If
    Next i
End Sub

Private Function GetStringValue(ByRef StringData() As Byte, ByVal StringLength As Integer, ByVal EncodingFormat As Byte) As String
    Dim i As Integer
    
    For i = 0 To StringLength - 1
        If EncodingFormat = 0 Or EncodingFormat = 3 Then
            'Clear text, null terminated.
            If StringData(i) = 0 Then Exit Function
            GetStringValue = GetStringValue & Chr$(StringData(i))
        ElseIf EncodingFormat = 1 Then
            'UNICODE text with BOM, double-null terminated.
            If i >= 2 And i Mod 2 = 0 Then
                If StringData(i) = 0 Then Exit Function
                GetStringValue = GetStringValue & Chr$(StringData(i))
            End If
        ElseIf EncodingFormat = 2 Then
            'UNICODE text without BOM, double-null terminated.
            If i Mod 2 = 0 Then
                If StringData(i) = 0 Then Exit Function
                GetStringValue = GetStringValue & Chr$(StringData(i))
            End If
        End If
        If Not EncodingFormat = 1 Or i >= 2 Then
        End If
    Next i
End Function

Private Sub SetStringValue(ByRef StringData() As Byte, ByVal Value As String, ByVal StringLength As Integer)
    Dim i As Integer
    
    For i = 0 To StringLength - 1
        StringData(i) = Asc(Mid$(Value, i + 1, 1))
    Next i
End Sub

'Bit Nr. 0 is the last bit, bit 7 the first bit.
Private Sub SetBit(ByVal BitNr As Integer, ByRef SrcData As Byte, ByVal BitState As Boolean)
    Dim Pattern As Byte
    
    If BitState Then
        'set a bit to 1
        Pattern = 2 ^ BitNr
        SrcData = SrcData Or Pattern
    Else
        'set a bit to 0
        Pattern = 255 - 2 ^ BitNr
        SrcData = SrcData And Pattern
    End If
End Sub

Private Function GetBit(ByVal BitNr As Byte, ByVal SrcData As Byte) As Boolean
    Dim Pattern As Byte
    
    Pattern = 2 ^ BitNr
    If SrcData And Pattern Then GetBit = True
End Function

