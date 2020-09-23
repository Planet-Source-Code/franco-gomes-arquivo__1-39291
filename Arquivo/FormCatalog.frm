VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormCatalog 
   Caption         =   "Catalog Disc Contents"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   Icon            =   "FormCatalog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   180
      Left            =   165
      TabIndex        =   7
      Top             =   1305
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton Com_ListFiles 
      Appearance      =   0  'Flat
      Caption         =   "&Go"
      Height          =   375
      Left            =   5175
      TabIndex        =   2
      Top             =   435
      Width           =   1245
   End
   Begin VB.CommandButton Com_Exit 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5175
      TabIndex        =   3
      Top             =   1110
      Width           =   1245
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   495
      Width           =   2340
   End
   Begin VB.TextBox Disc_Label 
      Height          =   315
      Left            =   2835
      TabIndex        =   1
      Top             =   495
      Width           =   2055
   End
   Begin VB.Label LabInfo 
      Height          =   225
      Left            =   165
      TabIndex        =   6
      Top             =   930
      Width           =   4710
   End
   Begin VB.Label Label1 
      Caption         =   "Disc Label"
      Height          =   285
      Left            =   2850
      TabIndex        =   5
      Top             =   255
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Select Drive"
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   255
      Width           =   1815
   End
End
Attribute VB_Name = "FormCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ListFiles() As String
Private LenghtDim As Long
Private FileCount As Long
Private Sub Com_ListFiles_Click()

    ListAll

End Sub
Private Sub Com_Exit_Click()

    Unload Me

End Sub
Sub ListAll()

On Error GoTo ListAll_Error

Dim Nome As String
Dim Disco As String
Dim Inicio As Long
Dim TotalRecords As Long
Dim RetBool As Boolean
Dim MyPath As String

    If Disc_Label.Text = vbNullString Then
        MsgBox "Enter some name, please.", vbInformation
        Disc_Label.SetFocus
        Exit Sub
    End If
    MyPath = App.Path
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If UCase(Dir(MyPath & "ArqFiles\" & Disc_Label.Text & ".arq")) = UCase(Disc_Label.Text & ".arq") Then
        If MsgBox("File """ & Disc_Label.Text & ".arq" & """ already exists. Overwrite?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Com_Exit.Enabled = False
    Disco = Drive1.Drive
    Disco = UCase(Left(Disco, 2)) & "\"
    Nome = Dir(Disco, vbArchive + vbDirectory + vbHidden + vbReadOnly + vbSystem)
    LabInfo.Caption = "Reading data"
    LabInfo.Refresh
    
    
    LenghtDim = 2000
    ReDim ListFiles(LenghtDim)
    FileCount = 0
    Do While ListRoot(Disco, Nome) = True
        Nome = Dir
    Loop
    Inicio = 1
    RetBool = True
    TotalRecords = FileCount
    ProgressBar.Value = 0
    Do While RetBool = True
        RetBool = ListMore(Inicio, TotalRecords)
        Inicio = TotalRecords + 1
        TotalRecords = FileCount
        If ProgressBar.Value < 81 Then
            ProgressBar.Value = ProgressBar.Value + 20
            Else: ProgressBar.Value = 0
        End If
    Loop
    LabInfo.Caption = vbNullString
    LabInfo.Refresh
    StoreData
    LabInfo.Caption = vbNullString
    LabInfo.Refresh
    Screen.MousePointer = vbNormal
    MsgBox "Writting data from [" & Disc_Label.Text & "] finished!", vbInformation
    Com_Exit.Enabled = True
    ProgressBar.Value = 0
    ReDim ListFiles(0)
    
Exit Sub
    
ListAll_Error:
    LabInfo.Caption = vbNullString
    LabInfo.Refresh
    Screen.MousePointer = vbNormal
    Com_Exit.Enabled = True
    ProgressBar.Value = 0
    MsgBox "Error reading " & Disco, vbCritical
    Err.Clear

End Sub
Private Function ListRoot(FDisc As String, FName As String) As Boolean

On Error GoTo ListRoot_Error

    ListRoot = False
    If FName = vbNullString Then Exit Function
    ListRoot = True
    If FName <> "." And FName <> ".." Then
        FileCount = FileCount + 1
        If FileCount > LenghtDim Then
            LenghtDim = LenghtDim + 2000
            ReDim Preserve ListFiles(LenghtDim)
        End If
        If UCase(FName) = "HIBERFIL.SYS" Or UCase(FName) = "PAGEFILE.SYS" Then
            ListFiles(FileCount) = FDisc & FName
            Else
            If (GetAttr(FDisc & FName) And vbDirectory) = vbDirectory Then
                ListFiles(FileCount) = "<" & FDisc & FName
                Else: ListFiles(FileCount) = FDisc & FName
            End If
        End If
    End If
    
Exit Function
    
ListRoot_Error:
    LabInfo.Caption = vbNullString
    LabInfo.Refresh
    Screen.MousePointer = vbNormal
    Com_Exit.Enabled = True
    ProgressBar.Value = 0
    MsgBox "Error reading " & FDisc, vbCritical
    Err.Clear

End Function
Private Function ListMore(IniPos As Long, LastPos As Long) As Boolean

On Error GoTo ListMore_Error

Dim Cb As Long
Dim Dirc As String
Dim Disk As String
Dim Nom As String

    For Cb = IniPos To LastPos
        Dirc = ListFiles(Cb)
        If Left(Dirc, 1) = "<" Then
            Disk = Right(Dirc, Len(Dirc) - 1) & "\"
            Nom = Dir(Disk, vbArchive + vbDirectory + vbHidden + vbReadOnly + vbSystem)
            Do Until Nom = vbNullString
                If Nom <> "." And Nom <> ".." Then
                    ListMore = True
                    FileCount = FileCount + 1
                    If FileCount > LenghtDim Then
                        LenghtDim = LenghtDim + 2000
                        ReDim Preserve ListFiles(LenghtDim)
                    End If
                    If (GetAttr(Disk & Nom) And vbDirectory) = vbDirectory Then
                        ListFiles(FileCount) = "<" & Disk & Nom
                        Else: ListFiles(FileCount) = Disk & Nom
                    End If
                End If
                Nom = Dir
            Loop
        End If
    Next Cb
    
Exit Function
    
ListMore_Error:
    LabInfo.Caption = vbNullString
    LabInfo.Refresh
    Screen.MousePointer = vbNormal
    Com_Exit.Enabled = True
    ProgressBar.Value = 0
    MsgBox "Error reading " & Disk, vbCritical
    Err.Clear

End Function
Sub StoreData()

On Error GoTo StoreData_Error

Dim Ca As Long
Dim Cb As Long
Dim Remover As Boolean
Dim Txt As String
Dim Fs As Object
Dim F As Object
Dim MyPath As String

    MyPath = App.Path
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If Dir(MyPath & "ArqFiles", vbDirectory + vbReadOnly + vbHidden) = vbNullString Then
        MkDir MyPath & "ArqFiles"
    End If
    LabInfo.Caption = "Writing data to file: " & MyPath & "\" & Disc_Label.Text & ".arq"
    LabInfo.Refresh
    ProgressBar.Value = 0
    
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set F = Fs.CreateTextFile(MyPath & "ArqFiles\" & Disc_Label.Text & ".arq", True)
    
    For Ca = 1 To FileCount
        Txt = ListFiles(Ca)
        If Left(Txt, 1) <> "<" Then
            Txt = Right(Txt, Len(Txt) - 3)
            F.WriteLine (Txt)
            Cb = Int((100 * Ca) / FileCount)
            If Cb > ProgressBar.Value Then
                ProgressBar.Value = Cb
                ProgressBar.Refresh
            End If
        End If
    Next Ca
    F.Close
    DoEvents
    Set Fs = Nothing
    Set F = Nothing
    
Exit Sub

StoreData_Error:
    MsgBox "Error writing file " & MyPath & "ArqFiles\" & Disc_Label.Text & ".arq", vbCritical
    Err.Clear

End Sub
Private Sub Form_Load()

Me.Icon = FormArquivo.Icon
    
ProgressBar.Value = 0

End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set FormCatalog = Nothing

End Sub
