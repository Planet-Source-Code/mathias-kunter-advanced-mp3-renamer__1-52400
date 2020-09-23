VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Advanced MP3 Renamer v 1.0"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraParams 
      Caption         =   "Renaming parameters for this file(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   5820
      TabIndex        =   13
      Top             =   60
      Width           =   5955
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   4740
         Width           =   4575
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   4380
         Width           =   4575
      End
      Begin VB.OptionButton optSource 
         Caption         =   "User entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   27
         Top             =   4080
         Width           =   1695
      End
      Begin VB.ComboBox cmbTitle 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   3540
         Width           =   3375
      End
      Begin VB.ComboBox cmbArtist 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   3120
         Width           =   3375
      End
      Begin VB.OptionButton optSource 
         Caption         =   "File name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   2460
         Width           =   1695
      End
      Begin VB.OptionButton optSource 
         Caption         =   "ID3v1 tag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   1380
         Width           =   1575
      End
      Begin VB.OptionButton optSource 
         Caption         =   "ID3v2 tag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         X1              =   120
         X2              =   5820
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Old file name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   5640
         Width           =   1185
      End
      Begin VB.Label lblOldFileName 
         AutoSize        =   -1  'True
         Caption         =   "<Old file name>"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1500
         TabIndex        =   33
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "New file name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   5940
         Width           =   1275
      End
      Begin VB.Label lblNewFileName 
         AutoSize        =   -1  'True
         Caption         =   "<New file name>"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1500
         TabIndex        =   31
         Top             =   5940
         Width           =   1185
      End
      Begin VB.Label lblUserArtistInfo 
         AutoSize        =   -1  'True
         Caption         =   "Artist:"
         Height          =   195
         Left            =   420
         TabIndex        =   29
         Top             =   4440
         Width           =   390
      End
      Begin VB.Label lblUserTitleInfo 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   420
         TabIndex        =   28
         Top             =   4800
         Width           =   345
      End
      Begin VB.Label lblFileTitleInfo 
         AutoSize        =   -1  'True
         Caption         =   "Interpet this item as title:"
         Height          =   195
         Left            =   420
         TabIndex        =   25
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblFileArtistInfo 
         AutoSize        =   -1  'True
         Caption         =   "Interpet this item as artist:"
         Height          =   195
         Left            =   420
         TabIndex        =   24
         Top             =   3180
         Width           =   1785
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         Caption         =   "<File name>"
         Height          =   195
         Left            =   1320
         TabIndex        =   23
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblFileNameInfo 
         AutoSize        =   -1  'True
         Caption         =   "File name:"
         Height          =   195
         Left            =   420
         TabIndex        =   22
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label lblv1Title 
         AutoSize        =   -1  'True
         Caption         =   "<Title>"
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label lblv1Artist 
         AutoSize        =   -1  'True
         Caption         =   "<Artist>"
         Height          =   195
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label lblv1TitleInfo 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   420
         TabIndex        =   19
         Top             =   1980
         Width           =   345
      End
      Begin VB.Label lblv1ArtistInfo 
         AutoSize        =   -1  'True
         Caption         =   "Artist:"
         Height          =   195
         Left            =   420
         TabIndex        =   18
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label lblv2ArtistInfo 
         AutoSize        =   -1  'True
         Caption         =   "Artist:"
         Height          =   195
         Left            =   420
         TabIndex        =   17
         Top             =   660
         Width           =   390
      End
      Begin VB.Label lblv2TitleInfo 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblv2Artist 
         AutoSize        =   -1  'True
         Caption         =   "<Artist>"
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   660
         Width           =   525
      End
      Begin VB.Label lblv2Title 
         AutoSize        =   -1  'True
         Caption         =   "<Title>"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   960
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog CommonDlg 
      Left            =   120
      Top             =   7500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   32767
   End
   Begin VB.Frame Frame1 
      Caption         =   "General renaming parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   5820
      TabIndex        =   26
      Top             =   6480
      Width           =   5955
      Begin VB.CheckBox chkRenameFile 
         Caption         =   "Rename files"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1080
         Value           =   1  'Aktiviert
         Width           =   1875
      End
      Begin VB.CheckBox chkRewriteIDv1 
         Caption         =   "Rewrite ID3v1 tags"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   780
         Value           =   1  'Aktiviert
         Width           =   1875
      End
      Begin VB.CheckBox chkRewriteIDv2 
         Caption         =   "Rewrite ID3v2 tags (slow)"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Aktiviert
         Width           =   2235
      End
      Begin VB.CheckBox chkUpperCase 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Upper case first letter of every word"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   780
         Value           =   1  'Aktiviert
         Width           =   3015
      End
      Begin VB.CheckBox chkLowerCase 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Lower case following letters of every word"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   1080
         Value           =   1  'Aktiviert
         Width           =   3495
      End
   End
   Begin VB.ListBox lstFiles 
      Height          =   7860
      Left            =   120
      MultiSelect     =   2  'Erweitert
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please add or select files..."
      Height          =   195
      Left            =   5820
      TabIndex        =   35
      Top             =   120
      Width           =   1890
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Ready."
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   8040
      Width           =   510
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuRename 
         Caption         =   "Rename MP3z now"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
         Begin VB.Menu mnuAddFile 
            Caption         =   "Files..."
         End
         Begin VB.Menu mnuAddDir 
            Caption         =   "Directories..."
         End
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Remove"
         Begin VB.Menu mnuClearSelect 
            Caption         =   "Selected items"
         End
         Begin VB.Menu mnuClearList 
            Caption         =   "Complete list"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bStop As Boolean


'******************************************MENU******************************************
Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuRename_Click()
    Dim i As Long
    Dim bFirst As Boolean
    Dim TopIndex As Long, bRemove() As Boolean, WritePos As Long
    Dim NewArtist As String, NewTitle As String, NewTag As ID3Tag
    
    If MsgBox("This is going to rename all files in the list. Also notice that the files mustn't be write-protected for tag rewriting. Continue?", vbYesNo, "Advanced MP3 Renamer") = vbNo Then Exit Sub
    If nFiles = 0 Then Exit Sub
    TopIndex = lstFiles.TopIndex
    bFirst = True
    ReDim bRemove(nFiles - 1)
    Me.MousePointer = 11
    Me.Enabled = False
    
    'Go for it!
    For i = 0 To nFiles - 1
        With RdFiles(i)
            If .SourceType = SOURCE_IDV2 Then
                NewArtist = CleanStr(.IDv2.Artist, chkUpperCase.Value, chkLowerCase.Value, False)
                NewTitle = CleanStr(.IDv2.Title, chkUpperCase.Value, chkLowerCase.Value, False)
                NewTag.Album = CleanStr(.IDv2.Album, chkUpperCase.Value, chkLowerCase.Value, False)
            ElseIf .SourceType = SOURCE_IDV1 Then
                NewArtist = CleanStr(.IDv1.Artist, chkUpperCase.Value, chkLowerCase.Value, False)
                NewTitle = CleanStr(.IDv1.Title, chkUpperCase.Value, chkLowerCase.Value, False)
                NewTag.Album = CleanStr(.IDv1.Album, chkUpperCase.Value, chkLowerCase.Value, False)
            ElseIf .SourceType = SOURCE_FILENAME Then
                NewArtist = CleanStr(.FileTag.Artist, chkUpperCase.Value, chkLowerCase.Value, False)
                NewTitle = CleanStr(.FileTag.Title, chkUpperCase.Value, chkLowerCase.Value, False)
            ElseIf .SourceType = SOURCE_USERENTRY Then
                NewArtist = CleanStr(.UserTag.Artist, False, False, False)
                NewTitle = CleanStr(.UserTag.Title, False, False, False)
            End If
            NewTag.Artist = NewArtist
            NewTag.Title = NewTitle
            'Delete every file from the list which could be edited successfully.
            bRemove(i) = True
            If NewArtist = "" Or NewTitle = "" Then
                bRemove(i) = False
            Else
                If chkRewriteIDv2.Value = 1 Then
                    'Rewrite IDv2 tag.
                    If Not .IDv2.Artist = NewTag.Artist Or Not .IDv2.Title = NewTag.Title Then
                        If Not mdlID3.WriteID3v2(.SourceFile, NewTag) Then bRemove(i) = False
                    End If
                End If
                If bRemove(i) And chkRewriteIDv1.Value = 1 Then
                    'Rewrite IDv1 tag.
                    If Not .IDv1.Artist = NewTag.Artist Or Not .IDv1.Title = NewTag.Title Then
                        If Not mdlID3.WriteID3v1(.SourceFile, NewTag) Then bRemove(i) = False
                    End If
                End If
                If bRemove(i) And chkRenameFile.Value = 1 Then
                    'Rename the file.
                    If MoveFile(.SourceFile, GetDir(.SourceFile) & NewArtist & " - " & NewTitle & ".mp3") = 0 Then bRemove(i) = False
                End If
            End If
            
            lblStatus.Caption = "Renaming files... " & Int(100 * ((i + 1) / nFiles)) & "% done"
            DoEvents
        End With
    Next i
    
    Me.Enabled = True
    Me.MousePointer = 0
    lblStatus.Caption = "Ready."
    
    'Remove every file from the list which could be successfully renamed, thus only keeping invalid items in the list.
    For i = 0 To nFiles - 1
        If Not bRemove(i) Then
            If bFirst Then
                MsgBox "There were some files in the list which couldn't be renamed because of missing renaming informations or identical file names." & vbNewLine & "These files were left in the list for further editing.", vbInformation, "Advanced MP3 Renamer"
                bFirst = False
            End If
            RdFiles(WritePos) = RdFiles(i)
            WritePos = WritePos + 1
        End If
    Next i
    nFiles = WritePos
    ShowFiles
    If TopIndex > lstFiles.ListCount - 1 Then TopIndex = lstFiles.ListCount - 1
    If Not TopIndex = -1 Then lstFiles.TopIndex = TopIndex
End Sub

Private Sub mnuAddFile_Click()
    Dim i As Long
    Dim sRdFiles() As String, sRdPath As String, lNrFiles As Long
    
    'Show the open form.
    CommonDlg.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    CommonDlg.DefaultExt = "*.mp3"
    CommonDlg.FileName = "*.mp3"
    CommonDlg.Filter = "MP3 files (*.mp3)|*.mp3"
    CommonDlg.ShowOpen
    If CommonDlg.FileName = "*.mp3" Or CommonDlg.FileName = "" Then Exit Sub
    
    'Get the files out of the returned string.
    sRdFiles = Split(CommonDlg.FileName, Chr$(0))
    lNrFiles = UBound(sRdFiles) + 1
    If lNrFiles > 1 Then
        'Multi-Select
        sRdPath = NormalizeDir(sRdFiles(0))
        For i = 0 To lNrFiles - 2
            sRdFiles(i) = sRdPath & sRdFiles(i + 1)
        Next i
        lNrFiles = lNrFiles - 1
    End If
    
    'Add the files.
    AddFiles sRdFiles, lNrFiles
    'Show the files.
    ShowFiles
End Sub

'Add a folder with mp3z in it.
Private Sub mnuAddDir_Click()
    Dim lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    Dim RdStrings() As String, nNewFiles As Long

    'Let the user select a folder with mp3z in it.
    'Show a browse-for-folder form:
    With udtBI
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat("Please select a folder with MP3z in it:", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList = 0 Then Exit Sub
        
    'Get the selected folder.
    sPath = String$(MAX_PATH, 0)
    SHGetPathFromIDList lpIDList, sPath
    CoTaskMemFree lpIDList
    sPath = StripNulls(sPath)
    
    'Search for mp3 files in the selected folder recursivly.
    lblStatus.Caption = "Searching for mp3 files..."
    If MsgBox("Do you want to include all subdirectories of this folder?", vbYesNo Or vbQuestion, "Advanced MP3 Renamer") = vbYes Then
        nNewFiles = FindFilesAPI(sPath, ".mp3", RdStrings(), True)
    Else
        nNewFiles = FindFilesAPI(sPath, ".mp3", RdStrings(), False)
    End If
    
    If nNewFiles = 0 Then
        MsgBox "There were no mp3 files found in this folder.", vbInformation, "Advanced MP3 Renamer"
    Else
        'Add the files.
        AddFiles RdStrings, nNewFiles
        'Show the files.
        ShowFiles
    End If
End Sub

Private Sub mnuClearSelect_Click()
    Dim i As Long
    Dim TopIndex As Long, bRemove() As Boolean, WritePos As Long
    
    If nFiles = 0 Then Exit Sub
    TopIndex = lstFiles.TopIndex
    ReDim bRemove(nFiles - 1)
    'First, mark all selected items.
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            bRemove(lstFiles.ItemData(i)) = True
        End If
    Next i
    'Then, remove them from the list.
    For i = 0 To nFiles - 1
        If Not bRemove(i) Then
            RdFiles(WritePos) = RdFiles(i)
            WritePos = WritePos + 1
        End If
    Next i
    nFiles = WritePos
    ShowFiles
    If TopIndex > lstFiles.ListCount - 1 Then TopIndex = lstFiles.ListCount - 1
    If Not TopIndex = -1 Then lstFiles.TopIndex = TopIndex
End Sub

Private Sub mnuClearList_Click()
    If MsgBox("Do you really want to clear the current list?", vbYesNo Or vbQuestion, "Advanced MP3 Renamer") = vbNo Then Exit Sub
    lstFiles.Clear
    nFiles = 0
    lstFiles_Click
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Advanced MP3 Renamer v 1.0" & vbNewLine & "Copyright by Mathias Kunter" & vbNewLine & "Mail: mathiaskunter@yahoo.de", vbInformation, "Advanced MP3 Renamer"
End Sub





'******************************************FORM******************************************

Private Sub Form_Load()
    'Refresh command bar.
    lstFiles_Click
    'Refresh menu.
    RefreshMenu
End Sub

Private Sub lstFiles_Click()
    Dim i As Integer, j As Integer
    Dim Index As Long, bIDv2All As Boolean, bIDv1All As Boolean
    Dim ArtistAll As Long, TitleAll As Long, SourceAll As MP3SourceEnum
    Dim SameItems() As String, SameItemCnt As Long
    Dim bState As Boolean, bFirst As Boolean
    
    If bStop Then Exit Sub
    bStop = True
    
    RefreshMenu
    
    If Not lstFiles.SelCount = 0 Then
        'Enable default commands.
        fraParams.Visible = True
        
        'Read out the settings which are to display from the selected items.
        bIDv2All = True
        bIDv1All = True
        bFirst = True
        For i = 0 To lstFiles.ListCount - 1
            If lstFiles.Selected(i) Then
                With RdFiles(lstFiles.ItemData(i))
                    If Not .HasIDv2 Then bIDv2All = False
                    If Not .HasIDv1 Then bIDv1All = False
                    If bFirst Then
                        ArtistAll = .FileInterpretArtist
                        TitleAll = .FileInterpretTitle
                        SourceAll = .SourceType
                        Index = lstFiles.ItemData(i)
                        SameItemCnt = .FileInterpretItemCnt
                        ReDim SameItems(SameItemCnt - 1)
                        For j = 0 To SameItemCnt - 1
                            SameItems(j) = .FileInterpretItems(j)
                        Next j
                        bFirst = False
                    Else
                        If Not .FileInterpretArtist = ArtistAll Then ArtistAll = -1
                        If Not .FileInterpretTitle = TitleAll Then TitleAll = -1
                        If Not .SourceType = SourceAll Then SourceAll = -1
                        If .FileInterpretItemCnt < SameItemCnt Then SameItemCnt = .FileInterpretItemCnt
                        For j = 0 To SameItemCnt - 1
                            If Not LCase$(Trim$(SameItems(j))) = LCase$(Trim$(.FileInterpretItems(j))) Then
                                SameItems(j) = ""
                            End If
                        Next j
                        Index = -1
                    End If
                End With
            End If
        Next i
        
        'Set the current data source state and copy the text into the user input text boxes and the final
        'renaming parameters.
        If Not Index = -1 Then
            lblOldFileName.Caption = GetFile(RdFiles(Index).SourceFile)
        Else
            lblOldFileName.Caption = "<Unable to display - multi select>"
        End If
        If SourceAll = -1 Then
            'Files use different renaming information sources.
            For i = 0 To 3
                optSource(i).Value = False
            Next i
            'Index must be -1 here becuase there are different source types.
            lblNewFileName.Caption = "<Unable to display - multi select>"
        ElseIf SourceAll = SOURCE_IDV2 Then
            optSource(0).Value = True
            If Not Index = -1 Then
                With RdFiles(Index)
                    txtArtist.Text = .IDv2.Artist
                    txtTitle.Text = .IDv2.Title
                    .UserTag = .IDv2
                    lblNewFileName.Caption = CleanStr(.IDv2.Artist, chkUpperCase.Value, chkLowerCase.Value, False) & " - " & CleanStr(.IDv2.Title, chkUpperCase.Value, chkLowerCase.Value, False) & ".mp3"
                End With
            End If
        ElseIf SourceAll = SOURCE_IDV1 Then
            optSource(1).Value = True
            If Not Index = -1 Then
                With RdFiles(Index)
                    txtArtist.Text = .IDv1.Artist
                    txtTitle.Text = .IDv1.Title
                    .UserTag = .IDv1
                    lblNewFileName.Caption = CleanStr(.IDv1.Artist, chkUpperCase.Value, chkLowerCase.Value, False) & " - " & CleanStr(.IDv1.Title, chkUpperCase.Value, chkLowerCase.Value, False) & ".mp3"
                End With
            End If
        ElseIf SourceAll = SOURCE_FILENAME Then
            optSource(2).Value = True
            If Not Index = -1 Then
                With RdFiles(Index)
                    txtArtist.Text = .FileTag.Artist
                    txtTitle.Text = .FileTag.Title
                    .UserTag = .FileTag
                    lblNewFileName.Caption = CleanStr(.FileTag.Artist, chkUpperCase.Value, chkLowerCase.Value, False) & " - " & CleanStr(.FileTag.Title, chkUpperCase.Value, chkLowerCase.Value, False) & ".mp3"
                End With
            End If
        ElseIf SourceAll = SOURCE_USERENTRY Then
            optSource(3).Value = True
            If Not Index = -1 Then
                With RdFiles(Index)
                    txtArtist.Text = .UserTag.Artist
                    txtTitle.Text = .UserTag.Title
                    lblNewFileName.Caption = CleanStr(.UserTag.Artist, False, False, False) & " - " & CleanStr(.UserTag.Title, False, False, False) & ".mp3"
                End With
            End If
        End If
        If Index = -1 Then
            txtArtist.Text = ""
            txtTitle.Text = ""
            lblNewFileName.Caption = "<Unable to display - multi select>"
        End If
        
        'Update data output.
        'IDv2 tag info
        If bIDv2All Then
            bState = True
            If Not Index = -1 Then
                lblv2Artist.Caption = RdFiles(Index).IDv2.Artist
                lblv2Title.Caption = RdFiles(Index).IDv2.Title
            Else
                lblv2Artist.Caption = "<Unable to display - multi select>"
                lblv2Title.Caption = "<Unable to display - multi select>"
            End If
        Else
            bState = False
            lblv2Artist.Caption = "<IDv2 tag is not present or invalid>"
            lblv2Title.Caption = "<IDv2 tag is not present or invalid>"
        End If
        optSource(0).Enabled = bState
        
        'IDv1 tag info
        If bIDv1All Then
            bState = True
            If Not Index = -1 Then
                lblv1Artist.Caption = RdFiles(Index).IDv1.Artist
                lblv1Title.Caption = RdFiles(Index).IDv1.Title
            Else
                lblv1Artist.Caption = "<Unable to display - multi select>"
                lblv1Title.Caption = "<Unable to display - multi select>"
            End If
        Else
            bState = False
            lblv1Artist.Caption = "<IDv1 tag is not present or invalid>"
            lblv1Title.Caption = "<IDv1 tag is not present or invalid>"
        End If
        optSource(1).Enabled = bState
        
        'File name info
        cmbArtist.Enabled = optSource(2).Value
        cmbTitle.Enabled = optSource(2).Value
        If Not Index = -1 Then
            lblFileName.Caption = GetFile(RdFiles(Index).SourceFile)
        Else
            lblFileName.Caption = "<Unable to display - multi select>"
        End If
        cmbArtist.Clear
        cmbArtist.AddItem "<N/A>"
        For i = 0 To SameItemCnt - 1
            If Not SameItems(i) = "" Then
                cmbArtist.AddItem i + 1 & ".: " & SameItems(i)
            Else
                cmbArtist.AddItem i + 1 & ".: <Not uniformly>"
            End If
        Next i
        cmbArtist.ListIndex = ArtistAll + 1
        cmbTitle.Clear
        cmbTitle.AddItem "<N/A>"
        For i = 0 To SameItemCnt - 1
            If Not SameItems(i) = "" Then
                cmbTitle.AddItem i + 1 & ".: " & SameItems(i)
            Else
                cmbTitle.AddItem i + 1 & ".: <Not uniformly>"
            End If
        Next i
        cmbTitle.ListIndex = TitleAll + 1
        
        'User input info
        If Index = -1 Then
            optSource(3).Enabled = False
            txtArtist.Text = ""
            txtTitle.Text = ""
        Else
            optSource(3).Enabled = True
        End If
        If Not optSource(3).Value Or Index = -1 Then
            bState = False
        Else
            bState = True
        End If
        txtArtist.Enabled = bState
        txtTitle.Enabled = bState
    Else
        'Nothing's selected, don't show the parameter frame.
        fraParams.Visible = False
    End If
    bStop = False
End Sub

Private Sub lstFiles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        'Entf key
        mnuClearSelect_Click
    End If
End Sub

Private Sub optSource_Click(Index As Integer)
    Dim i As Long
    
    If bStop Then Exit Sub
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            With RdFiles(lstFiles.ItemData(i))
                If Index = 0 Then
                    .SourceType = SOURCE_IDV2
                ElseIf Index = 1 Then
                    .SourceType = SOURCE_IDV1
                ElseIf Index = 2 Then
                    .SourceType = SOURCE_FILENAME
                ElseIf Index = 3 Then
                    .SourceType = SOURCE_USERENTRY
                End If
            End With
        End If
    Next i
    lstFiles_Click
End Sub

Private Sub cmbArtist_Click()
    Dim i As Long
    
    If bStop Then Exit Sub
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            With RdFiles(lstFiles.ItemData(i))
                .FileInterpretArtist = cmbArtist.ListIndex - 1
                If Not .FileInterpretArtist = -1 Then
                    .FileTag.Artist = .FileInterpretItems(.FileInterpretArtist)
                Else
                    .FileTag.Artist = ""
                End If
            End With
        End If
    Next i
    lstFiles_Click
End Sub

Private Sub cmbTitle_Click()
    Dim i As Long
    
    If bStop Then Exit Sub
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            With RdFiles(lstFiles.ItemData(i))
                .FileInterpretTitle = cmbTitle.ListIndex - 1
                If Not .FileInterpretTitle = -1 Then
                    .FileTag.Title = .FileInterpretItems(.FileInterpretTitle)
                Else
                    .FileTag.Title = ""
                End If
            End With
        End If
    Next i
    lstFiles_Click
End Sub

Private Sub txtArtist_Change()
    Dim i As Long
    
    If bStop Then Exit Sub
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            With RdFiles(lstFiles.ItemData(i))
                .UserTag.Artist = txtArtist.Text
            End With
            Exit For
        End If
    Next i
    lstFiles_Click
End Sub

Private Sub txtTitle_Change()
    Dim i As Long
    
    If bStop Then Exit Sub
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            With RdFiles(lstFiles.ItemData(i))
                .UserTag.Title = txtTitle.Text
            End With
            Exit For
        End If
    Next i
    lstFiles_Click
End Sub

Private Sub chkUpperCase_Click()
    If bStop Then Exit Sub
    lstFiles_Click
End Sub

Private Sub chkLowerCase_Click()
    If bStop Then Exit Sub
    lstFiles_Click
End Sub

Private Sub chkRewriteIDv2_Click()
    RefreshMenu
End Sub

Private Sub chkRewriteIDv1_Click()
    RefreshMenu
End Sub

Private Sub chkRenameFile_Click()
    RefreshMenu
End Sub








'******************************************ALLGEMEIN******************************************

Private Sub RefreshMenu()
    If lstFiles.ListCount = 0 Then
        mnuClearList.Enabled = False
    Else
        mnuClearList.Enabled = True
    End If
    If lstFiles.ListCount = 0 Or (chkRewriteIDv2.Value = 0 And chkRewriteIDv1.Value = 0 And chkRenameFile.Value = 0) Then
        mnuRename.Enabled = False
    Else
        mnuRename.Enabled = True
    End If
    If lstFiles.SelCount = 0 Then
        mnuClearSelect.Enabled = False
    Else
        mnuClearSelect.Enabled = True
    End If
End Sub


Private Sub AddFiles(ByRef FileList() As String, ByVal FileCnt As Long)
    Dim i As Long
    
    'Add the new files to the track list.
    If FileCnt = 0 Then Exit Sub
    ReDim Preserve RdFiles(nFiles + FileCnt - 1)
    
    Me.MousePointer = 11
    Me.Enabled = False
    
    For i = 0 To FileCnt - 1
        RdFiles(nFiles + i) = AnalyseFile(FileList(i))
        lblStatus.Caption = "Reading files... " & Int(100 * ((i + 1) / FileCnt)) & "% done"
        DoEvents
    Next i
    nFiles = nFiles + FileCnt
    
    Me.Enabled = True
    Me.MousePointer = 0
    lblStatus.Caption = "Ready."
End Sub

Private Function AnalyseFile(ByVal strFile As String) As MP3File
    Dim SplitPos As Long
    
    With AnalyseFile
        .SourceFile = strFile
        
        
        '***Read and process IDv2 tag***
        .HasIDv2 = ReadID3v2(.SourceFile, .IDv2)
        If .HasIDv2 Then
            .IDv2.Artist = CleanInterpreteItems(.IDv2.Artist)
            .IDv2.Title = CleanInterpreteItems(.IDv2.Title)
            .IDv2.Album = CleanInterpreteItems(.IDv2.Album)
            If .IDv2.Title = "" And .IDv2.Artist = "" Then
                .HasIDv2 = False
            ElseIf Not .IDv2.Title = "" And .IDv2.Artist = "" Then
                'Only valid if in the title item is also the artist stored.
                SplitPos = InStr(1, .IDv2.Title, "-", vbTextCompare)
                If Not SplitPos = 0 Then
                    .IDv2.Artist = Trim$(Left$(.IDv2.Title, SplitPos - 1))
                    .IDv2.Title = Trim$(Right$(.IDv2.Title, Len(.IDv2.Title) - SplitPos))
                Else
                    'Set the tag to invalid.
                    .HasIDv2 = False
                End If
            End If
        End If
        
        
        '***Read and process IDv1 tag***
        .HasIDv1 = ReadID3v1(.SourceFile, .IDv1)
        If .HasIDv1 Then
            .IDv1.Artist = CleanInterpreteItems(.IDv1.Artist)
            .IDv1.Title = CleanInterpreteItems(.IDv1.Title)
            .IDv1.Album = CleanInterpreteItems(.IDv1.Album)
            If .IDv1.Title = "" And .IDv1.Artist = "" Then
                .HasIDv1 = False
            ElseIf Not .IDv1.Title = "" And .IDv1.Artist = "" Then
                'Only valid if in the title item is also the artist stored.
                SplitPos = InStr(1, .IDv1.Title, "-", vbTextCompare)
                If Not SplitPos = 0 Then
                    .IDv1.Artist = Trim$(Left$(.IDv1.Title, SplitPos - 1))
                    .IDv1.Title = Trim$(Right$(.IDv1.Title, Len(.IDv1.Title) - SplitPos))
                Else
                    'Set the tag to invalid.
                    .HasIDv1 = False
                End If
            End If
        End If
        
        
        '***Read and process file names***
        '1) Remove any invalid parts in the file name which can be found out. This includes
        'numeric parts (=track numbers), invalid signs, multiple spaces and so on.
        .FileInterpretItemCnt = SplitInterpreteItems(GetFileWOExt(.SourceFile), False, False, .FileInterpretItems)
        
        '2) Try to get the artist out of the file name. MP3Renamer uses the first part of the file name, this is until the
        'first "-" appears in the string.
        If .FileInterpretItemCnt > 0 Then
            .FileTag.Artist = .FileInterpretItems(0)
            .FileInterpretArtist = 0
        Else
            .FileInterpretArtist = -1
        End If
        
        '3) Try to get out the title of the file name. MP3Renamer uses the last part of the file.
        If .FileInterpretItemCnt > 1 Then
            'Use the last part of the file name as track title.
            .FileTag.Title = .FileInterpretItems(.FileInterpretItemCnt - 1)
            .FileInterpretTitle = .FileInterpretItemCnt - 1
        Else
            .FileInterpretTitle = -1
        End If
        
        If .HasIDv2 Then
            .SourceType = SOURCE_IDV2
        ElseIf .HasIDv1 Then
            .SourceType = SOURCE_IDV1
        Else
            .SourceType = SOURCE_FILENAME
        End If
    End With
End Function


Private Sub ShowFiles()
    Dim i As Long
    
    lstFiles.Clear
    For i = 0 To nFiles - 1
        lstFiles.AddItem GetFile(RdFiles(i).SourceFile)
        lstFiles.ItemData(lstFiles.NewIndex) = i
    Next i
    lstFiles_Click
End Sub
