Attribute VB_Name = "mdlMP3Info"
'Private i As Integer
'Private strEmptyString As String
'Private rdbyte As Byte
'Private s As String

Public Type MP3Info
    BITRATE As Integer
    CHANNELS As String
    COPYRIGHT As Boolean
    CRC As Boolean
    EMPHASIS As String
    FREQ As Long
    LAYER As Integer
    LENGTH As Long
    MPEG As String
    ORIGINAL As Boolean
    SIZE As Long
End Type

Public Type ID3v1Tag
    id As String * 3
    Title As String * 30
    Artist As String * 30
    Album As String * 30
    Year As String * 4
    Comment As String * 30
    Genre As Byte
End Type

Public Type ID3v2Tag
    Version As String
    Title As String
    Artist As String
    Album As String
    Year As String
    Comment As String
    Genre As String
    Track As String
End Type

Private Version As Byte

Public Function GetID3v1Tag(ByVal strFileName As String, v1Tag As ID3v1Tag) As Boolean
    Dim lngFilesize As Long
    Dim fn As Integer
    
    If Dir(strFileName) = "" Then
        GetID3v1Tag = False
        Exit Function
    End If

    On Local Error GoTo errorhandler
    
    'Open the file
    fn = FreeFile
    Open strFileName For Binary As #fn                      'Open the file so we can read it
    lngFilesize = LOF(fn)                                   'Size of the file, in bytes

    'Check for an ID3v1 tag
    Get #fn, lngFilesize - 127, v1Tag.id
    
    If v1Tag.id = "TAG" Then 'If "TAG" is present, then we have a valid ID3v1 tag and will extract all available ID3v1 info from the file
        GetID3v1Tag = True
        Get #fn, , v1Tag.Title   'Always limited to 30 characters
        Get #fn, , v1Tag.Artist  'Always limited to 30 characters
        Get #fn, , v1Tag.Album   'Always limited to 30 characters
        Get #fn, , v1Tag.Year    'Always limited to 4 characters
        Get #fn, , v1Tag.Comment 'Always limited to 30 characters
        Get #fn, , v1Tag.Genre   'Always limited to 1 byte (?)
        
        'Repair v1Tag variables (ie remove chr(0))
        If InStr(1, v1Tag.Title, Chr(0), vbTextCompare) > 0 Then
            v1Tag.Title = Left(v1Tag.Title, InStr(1, v1Tag.Title, Chr(0), vbTextCompare) - 1)
        End If
        If InStr(1, v1Tag.Comment, Chr(0), vbTextCompare) > 0 Then
            v1Tag.Comment = Left(v1Tag.Comment, InStr(1, v1Tag.Comment, Chr(0), vbTextCompare) - 1)
        End If
        If InStr(1, v1Tag.Album, Chr(0), vbTextCompare) > 0 Then
            v1Tag.Album = Left(v1Tag.Album, InStr(1, v1Tag.Album, Chr(0), vbTextCompare) - 1)
        End If
        If InStr(1, v1Tag.Artist, Chr(0), vbTextCompare) > 0 Then
            v1Tag.Artist = Left(v1Tag.Artist, InStr(1, v1Tag.Artist, Chr(0), vbTextCompare) - 1)
        End If
        If InStr(1, v1Tag.Year, Chr(0), vbTextCompare) > 0 Then
            v1Tag.Year = Left(v1Tag.Year, InStr(1, v1Tag.Year, Chr(0), vbTextCompare) - 1)
        End If
    Else
        'No Tag information present
        v1Tag.id = ""
        GetID3v1Tag = False
    End If
        
    'Close the file
    Close
    Exit Function
        
errorhandler:
    Close
End Function


Public Function WriteID3v1Tag(ByVal strFileName As String, v1Tag As ID3v1Tag) As Boolean
    Dim fn As Integer
    
    fn = FreeFile
    
    With v1Tag
        Open strFileName For Binary Access Write As #i
            Seek #i, FileLen(FileName) - 127
            Put #i, , .id
            Put #i, , .Title
            Put #i, , .Artist
            Put #i, , .Album
            Put #i, , .Year
            Put #i, , .Comment
            Put #i, , .Genre
        Close #fn
    End With
End Function

Public Function GetID3v2Tag(ByVal strFileName As String, v2Tag As ID3v2Tag) As Boolean
    Dim lngFilesize As Long
    Dim fn As Integer
    Dim RdByte As Byte
    Dim lngHeaderPosition As Long
    Dim Tag2 As String
    
    Dim TitleField As String
    Dim ArtistField As String
    Dim AlbumField As String
    Dim YearField As String
    Dim GenreField As String
    Dim FieldSize As Long
    Dim SizeOffset As Long
    Dim FieldOffset As Long
    Dim TrackNbr As String
    Dim SituationField As String
    Dim CommentField As String
    
    
    On Local Error GoTo errorhandler

    If Dir(strFileName) = "" Then
        GetID3v2Tag = False
        Exit Function
    End If

    'Open the file
    fn = FreeFile
    Open strFileName For Binary As #fn                      'Open the file so we can read it
    lngFilesize = LOF(fn)                                   'Size of the file, in bytes

    'Check for a Header
            
    Get #fn, 1, RdByte
    lngHeaderPosition = 1
    Get #fn, 2, RdByte
    If (RdByte < 250 Or RdByte > 251) Then
        'We have an ID3v2 tag
        GetID3v2Tag = True
        If RdByte = 68 Then
            Get #fn, 3, RdByte
            If RdByte = 51 Then
                Dim R As Double
                Get #fn, 4, Version
                Get #fn, 7, RdByte
                R = RdByte * 20917152
                Get #fn, 8, RdByte
                R = R + (RdByte * 16384)
                Get #fn, 9, RdByte
                R = R + (RdByte * 128)
                Get #fn, 10, RdByte
                R = R + RdByte
                If R > lngFilesize Or R > 2147483647 Then
                    Exit Function
                End If
                Tag2 = Space$(R)
                Get #fn, 11, Tag2
                lngHeaderPosition = R + 11
            End If
        End If
    End If
    If Trim$(Tag2) = "" Then
        'ID3v2 tag is missing
        GetID3v2Tag = False
        Exit Function
    End If
   
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Determine if the ID3v2 tag is ID3v2.2 or ID3v2.3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Notes: I haven't tested reading an MP3 file that has a ID3v2.2 tag
   
'set version for external reading
    v2Tag.Version = Version
   
    Select Case Version
    
        Case 2 'ID3v2.2
        'Set the fieldnames for version 2.0
            TitleField = "TT2"
            ArtistField = "TOA"
            AlbumField = "TAL"
            YearField = "TYE"
            GenreField = "TCO"
            FieldOffset = 7
            SizeOffset = 5
            TrackNbr = "TRCK"
            CommentField = "COM"
       
        Case 3 'ID3v2.3
        'Set the fieldnames for version 3.0
            TitleField = "TIT2"
            ArtistField = "TPE1"
            AlbumField = "TALB"
            YearField = "TYER"
            GenreField = "TCON"
            TrackNbr = "TRCK"
            CommentField = "COMM"
       
            FieldOffset = 11
            SizeOffset = 7
        Case Else
        'We don't have a valid ID3v2 tag, so bail
            Exit Function
            
    End Select
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract track title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       i = InStr(Tag2, TitleField)
       If i > 0 Then
          'read the title
          FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
          If Version = 3 Then
             'check for compressed or encrypted field
             RdByte = Asc(Mid$(Tag2, i + 9))
             If (RdByte And 128) = True Or (RdByte And 64) = True Then GoTo ReadAlbum
          End If
          v2Tag.Title = Mid$(Tag2, i + FieldOffset, FieldSize)
        Else
        v2Tag.Title = ""
       End If
       
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract album title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadAlbum:
    i = InStr(Tag2, AlbumField)
    If i > 0 Then
       FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
       If Version = 3 Then
          'check for compressed or encrypted field
          RdByte = Asc(Mid$(Tag2, i + 9))
          If (RdByte And 128) = 128 Or (RdByte And 64) = 64 Then GoTo ReadArtist
       End If
       v2Tag.Album = Mid$(Tag2, i + FieldOffset, FieldSize)
    Else
        v2Tag.Album = ""
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract artist name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadArtist:
   i = InStr(Tag2, ArtistField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         RdByte = Asc(Mid$(Tag2, i + 9))
         If (RdByte And 128) = 128 Or (RdByte And 64) = 64 Then GoTo ReadYear
      End If
      v2Tag.Artist = Mid$(Tag2, i + FieldOffset, FieldSize)
    Else
        v2Tag.Artist = ""
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract year title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadYear:
   i = InStr(Tag2, YearField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         RdByte = Asc(Mid$(Tag2, i + 9))
         If (RdByte And 128) = 128 Or (RdByte And 64) = 64 Then GoTo ReadGenre
      End If
      v2Tag.Year = Mid$(Tag2, i + FieldOffset, FieldSize)
    Else
        v2Tag.Year = 0
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract genre
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadGenre:
   i = InStr(Tag2, GenreField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      
      If Version = 3 Then
         'check for compressed or encrypted field
         RdByte = Asc(Mid$(Tag2, i + 9))
         If (RdByte And 128) = 128 Or (RdByte And 64) = 64 Then GoTo ReadTrackNbr
      End If
      
      s = Mid$(Tag2, i + FieldOffset, FieldSize)
      If Left$(s, 1) = "(" Then
        v2Tag.Genre = Val(Mid$(s, 2, 2))
        
      Else
         v2Tag.Genre = i
         
         If i > 0 Then
            v2Tag.Genre = Int(i / 22)
         End If
      End If
    Else
        v2Tag.Genre = 0
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract track number
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadTrackNbr:
   i = InStr(Tag2, TrackNbr)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         RdByte = Asc(Mid$(Tag2, i + 9))
         If (RdByte And 128) = 128 Or (RdByte And 64) = 64 Then GoTo Done
      End If
      v2Tag.Track = Mid$(Tag2, i + FieldOffset, FieldSize)
    Else
        v2Tag.Track = 0
   End If
   

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get v2 Comment Information
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadV2Comment:
    'Currently Not Working
    'v2Tag.Comment = "Comment for V2 Currently Not Available"
    i = InStr(Tag2, CommentField)
    If i > 0 Then
        'Get Comment Information
        FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
        If Version = 3 Then
           'check for compressed or encrypted field
           RdByte = Asc(Mid$(Tag2, i + 13))
           If (RdByte And 128) = 128 Or (RdByte And 64) = 64 Then GoTo Done
        End If
        v2Tag.Comment = Mid$(Tag2, i + FieldOffset + 4, FieldSize - 4)
    Else
        v2Tag.Comment = ""
    End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We're done looking for ID3v2 info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Done:

'Fix any chr(0) errors
    If InStr(1, v2Tag.Title, Chr(0), vbTextCompare) > 0 Then
        v2Tag.Title = Left(v2Tag.Title, InStr(1, v2Tag.Title, Chr(0), vbTextCompare) - 1)
    End If
    If InStr(1, v2Tag.Comment, Chr(0), vbTextCompare) > 0 Then
        v2Tag.Comment = Left(v2Tag.Comment, InStr(1, v2Tag.Comment, Chr(0), vbTextCompare) - 1)
    End If
    If InStr(1, v2Tag.Album, Chr(0), vbTextCompare) > 0 Then
        v2Tag.Album = Left(v2Tag.Album, InStr(1, v2Tag.Album, Chr(0), vbTextCompare) - 1)
    End If
    If InStr(1, v2Tag.Artist, Chr(0), vbTextCompare) > 0 Then
        v2Tag.Artist = Left(v2Tag.Artist, InStr(1, v2Tag.Artist, Chr(0), vbTextCompare) - 1)
    End If
    If InStr(1, v2Tag.Year, Chr(0), vbTextCompare) > 0 Then
        v2Tag.Year = Left(v2Tag.Year, InStr(1, v2Tag.Year, Chr(0), vbTextCompare) - 1)
    End If
    If InStr(1, v2Tag.Track, Chr(0), vbTextCompare) > 0 Then
        v2Tag.Track = Left(v2Tag.Track, InStr(1, v2Tag.Track, Chr(0), vbTextCompare) - 1)
    End If


   Close
   
   Exit Function

errorhandler:
    Err.Clear
    GetID3v2Tag = False
End Function
