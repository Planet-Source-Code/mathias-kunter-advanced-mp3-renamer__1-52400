VERSION 5.00
Begin VB.Form frmResults 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MP3 Renamer"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkCase 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Case sensitive renaming"
      Height          =   195
      Left            =   9480
      TabIndex        =   20
      Top             =   180
      Width           =   2235
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename mp3z!"
      Height          =   555
      Left            =   8100
      TabIndex        =   15
      Top             =   7200
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   9960
      TabIndex        =   14
      Top             =   7200
      Width           =   1755
   End
   Begin VB.Frame Frame4 
      Caption         =   "Files which couldn't be analysed"
      Height          =   3435
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   5655
      Begin VB.ListBox lstUnknown 
         Height          =   2400
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   840
         Width           =   5235
      End
      Begin VB.Label Label5 
         Caption         =   "Note: You have to set the artist and track data manually if you want to rename this files."
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   5280
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Files which may should be renamed"
      Height          =   3435
      Left            =   6060
      TabIndex        =   8
      Top             =   660
      Width           =   5655
      Begin VB.ListBox lstName 
         Height          =   2400
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label4 
         Caption         =   "Note: You have to set the artist and track data manually if you want to rename this files."
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   5040
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Renaming informations"
      Height          =   2595
      Left            =   6060
      TabIndex        =   3
      Top             =   4320
      Width           =   5655
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3660
         TabIndex        =   18
         Top             =   1980
         Width           =   1575
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   780
         Width           =   4035
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   1200
         Width           =   4035
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         Caption         =   "<select a file>"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   420
         Width           =   975
      End
      Begin VB.Label lblAlbum 
         AutoSize        =   -1  'True
         Caption         =   "<select a file>"
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Album:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "New artist:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "New title:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1260
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files which can automatically be renamed"
      Height          =   3435
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   5655
      Begin VB.ListBox lstTag 
         Height          =   2790
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   5235
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data analysing results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCurFile As Integer
Private mCase As Boolean

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub

Private Sub chkCase_Click()
    If chkCase.Value = 1 Then
        mCase = True
    Else
        mCase = False
    End If
    DisplayResults
End Sub

Public Sub DisplayResults()
    Dim i As Integer
    Dim bRename As Boolean
    
    'Show the user the results of the analysis.
    lstTag.Clear
    lstName.Clear
    lstUnknown.Clear
    For i = 0 To nFiles - 1
        With RdFiles(i)
            bRename = False
            If mCase Then
                If Not GetFile(.SourceFile) = .Artist & " - " & .Title & ".mp3" Then bRename = True
            Else
                If Not LCase$(GetFile(.SourceFile)) = LCase(.Artist & " - " & .Title & ".mp3") Then bRename = True
            End If
            If bRename Then
                'This file could be renamed.
                If .GotName Then
                    'Both artist and title are well known.
                    lstTag.AddItem GetFileWOExt(.SourceFile)
                    lstTag.ItemData(lstTag.NewIndex) = i
                ElseIf .GotArtistUnsure And .GotTitleUnsure Then
                    'Both artist and title are maybe known.
                    lstName.AddItem GetFileWOExt(.SourceFile)
                    lstName.ItemData(lstName.NewIndex) = i
                Else
                    'At least the artist or the title is completly unknown.
                    lstUnknown.AddItem GetFileWOExt(.SourceFile)
                    lstUnknown.ItemData(lstUnknown.NewIndex) = i
                End If
            End If
        End With
    Next i
    Me.SetFocus
    mCurFile = -1
    DisplayCurFile
End Sub

Private Sub DisplayCurFile()
    If mCurFile = -1 Then
        lblFile.Caption = "<select a file>"
        txtArtist.Text = ""
        txtTitle.Text = ""
        lblAlbum.Caption = "<select a file>"
        cmdApply.Enabled = False
    Else
        With RdFiles(mCurFile)
            lblFile.Caption = GetFile(.SourceFile)
            txtArtist.Text = .Artist
            txtTitle.Text = .Title
            If Not .Album = "" Then
                lblAlbum.Caption = .Album
            Else
                lblAlbum.Caption = "<unknown>"
            End If
        End With
        cmdApply.Enabled = True
    End If
End Sub

Private Sub lstTag_Click()
    If lstTag.ListIndex = -1 Then Exit Sub
    mCurFile = lstTag.ItemData(lstTag.ListIndex)
    DisplayCurFile
End Sub

Private Sub lstTag_GotFocus()
    Dim i As Integer
    
    For i = 0 To lstName.ListCount - 1
        lstName.Selected(i) = False
    Next i
    For i = 0 To lstUnknown.ListCount - 1
        lstUnknown.Selected(i) = False
    Next i
End Sub

Private Sub lstName_Click()
    If lstName.ListIndex = -1 Then Exit Sub
    mCurFile = lstName.ItemData(lstName.ListIndex)
    DisplayCurFile
End Sub

Private Sub lstName_GotFocus()
    Dim i As Integer
    
    For i = 0 To lstTag.ListCount - 1
        lstTag.Selected(i) = False
    Next i
    For i = 0 To lstUnknown.ListCount - 1
        lstUnknown.Selected(i) = False
    Next i
End Sub

Private Sub lstUnknown_Click()
    If lstUnknown.ListIndex = -1 Then Exit Sub
    mCurFile = lstUnknown.ItemData(lstUnknown.ListIndex)
    DisplayCurFile
End Sub

Private Sub lstUnknown_GotFocus()
    Dim i As Integer
    
    For i = 0 To lstTag.ListCount - 1
        lstTag.Selected(i) = False
    Next i
    For i = 0 To lstName.ListCount - 1
        lstName.Selected(i) = False
    Next i
End Sub

Private Sub cmdApply_Click()
    If mCurFile = -1 Then
        cmdApply.Enabled = False
        Exit Sub
    End If
    With RdFiles(mCurFile)
        .Artist = Trim$(txtArtist.Text)
        .Title = Trim$(txtTitle.Text)
        .GotName = True
        DisplayResults
    End With
End Sub

Private Sub cmdRename_Click()
    Dim i As Integer
    Dim OldName As String, NewName As String
    
    'Well, go for it.
    Me.MousePointer = 11
    On Local Error Resume Next
    For i = 0 To lstTag.ListCount - 1
        With RdFiles(lstTag.ItemData(i))
            OldName = .SourceFile
            NewName = GetDir(.SourceFile) & .Artist & " - " & .Title & ".mp3"
            Name OldName As NewName
        End With
    Next i
    Me.MousePointer = 0
    Err.Clear
    MsgBox "Files have been renamed!", vbInformation
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
