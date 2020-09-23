VERSION 5.00
Begin VB.Form FormFind 
   Caption         =   "Find File/Folder"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox ListLabels 
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.ListBox ListNames 
      Height          =   1620
      ItemData        =   "FormFind.frx":0000
      Left            =   150
      List            =   "FormFind.frx":0002
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1245
      Width           =   5100
   End
   Begin VB.CommandButton ComFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   3675
      TabIndex        =   1
      Top             =   285
      Width           =   1575
   End
   Begin VB.CommandButton ComClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3675
      TabIndex        =   2
      Top             =   780
      Width           =   1575
   End
   Begin VB.TextBox SearchText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   540
      Width           =   3180
   End
   Begin VB.Label LabResult 
      Caption         =   "LabResult"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   495
      TabIndex        =   8
      Top             =   930
      Width           =   2865
   End
   Begin VB.Label LabLabelDisc 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2370
      TabIndex        =   7
      Top             =   2955
      Width           =   2880
   End
   Begin VB.Label Label3 
      Caption         =   "Searching string"
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   270
      Width           =   2550
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Disc Label:"
      Height          =   240
      Left            =   1125
      TabIndex        =   4
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000C000&
      Height          =   285
      Left            =   135
      Shape           =   3  'Circle
      Top             =   885
      Width           =   300
   End
End
Attribute VB_Name = "FormFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary
Private Fs As Object
Private F As Object
Private Ts As Object
Private Sub ComClose_Click()

    Unload Me

End Sub
Private Sub ComFind_Click()

    If SearchText.Text = vbNullString Then
        MsgBox "Enter some searching text, please.", vbInformation
        SearchText.SetFocus
        Exit Sub
    End If
    ComFind.Enabled = False
    Ball.BackColor = &HFF&
    Ball.Visible = True
    Ball.Refresh
    LabLabelDisc.Caption = vbNullString
    LabResult.Caption = vbNullString
    LabResult.Refresh
    SearchFile

End Sub
Private Sub Form_Load()

    Me.Icon = FormArquivo.Icon
    Ball.Visible = False
    LabResult.Caption = vbNullString

End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set Fs = Nothing
    Set F = Nothing
    Set Ts = Nothing
    Set FormFind = Nothing

End Sub
Private Sub SearchFile()

Dim FileName As String
Dim FilePath As String
Dim LineText As String
Dim SrchText As String

    ListNames.Clear
    ListLabels.Clear
    SrchText = UCase(SearchText.Text)
    FilePath = App.Path
    If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    FilePath = FilePath & "ArqFiles\"
    FileName = Dir(FilePath, vbArchive + vbReadOnly + vbHidden)
    Do While FileName <> vbNullString
        If UCase(Right(FileName, 4)) = ".ARQ" Then
            Set Fs = CreateObject("Scripting.FileSystemObject")
            Set F = Fs.GetFile(FilePath & FileName)
            Set Ts = F.OpenAsTextStream(1, -2)
            Do While Ts.AtEndOfStream <> True
                LineText = Ts.Readline
                If InStr(UCase(LineText), SrchText) > 0 Then
                    ListNames.AddItem LineText
                    ListLabels.AddItem Left(FileName, Len(FileName) - 4)
                End If
            Loop
            Ts.Close
        End If
        FileName = Dir
    Loop
    If ListNames.ListCount <> 0 Then ListNames.ListIndex = 0
    Ball.BackColor = &HC000&
    Ball.Refresh
    ComFind.Enabled = True
    LabResult.Caption = ListNames.ListCount & " items found."
    LabResult.Refresh

End Sub
Private Sub ListNames_Click()

    If ListNames.ListCount <> 0 Then
        LabLabelDisc.Caption = ListLabels.List(ListNames.ListIndex)
    End If

End Sub
