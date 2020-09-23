VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormPrint
   Caption         =   "Print Disc Contents"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4785
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox ListFolders
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   1275
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.ListBox ListNames
      Height          =   255
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1575
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton ComPrinterSetup
      Caption         =   "Printer &Setup"
      Height          =   375
      Left            =   3030
      TabIndex        =   3
      Top             =   180
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog Dialog
      Left            =   2295
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ComClose
      Caption         =   "&Close"
      Height          =   375
      Left            =   3030
      TabIndex        =   1
      Top             =   1305
      Width           =   1575
   End
   Begin VB.CommandButton ComPrint
      Caption         =   "&Print"
      Height          =   375
      Left            =   3030
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.ListBox ListFiles
      Height          =   1230
      Left            =   165
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   450
      Width           =   2580
   End
   Begin VB.Label Label1
      Caption         =   "Select file to print report"
      Height          =   255
      Left            =   165
      TabIndex        =   4
      Top             =   150
      Width           =   2595
   End
End
Attribute VB_Name = "FormPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ComClose_Click()

    Unload Me

End Sub
Private Sub ComPrint_Click()

    If ListFiles.ListCount = 0 Then
        MsgBox "Nothing to print.", vbInformation
        Exit Sub
    End If
    If MsgBox("Print report of """ & ListFiles.List(ListFiles.ListIndex) & """?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    GetPrintingData
    PrintReport
    Screen.MousePointer = vbNormal

End Sub
Private Sub ComPrinterSetup_Click()

    With Dialog
        .Flags = cdlPDPrintSetup + cdlPDNoWarning
        .ShowPrinter
    End With
    Me.Caption = "Printer: " & Printer.DeviceName

End Sub
Private Sub Form_Load()

Dim FileName As String
Dim FilePath As String

    Me.Icon = FormArquivo.Icon
    Me.Caption = "Printer: " & Printer.DeviceName
    FilePath = App.Path
    If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    FilePath = FilePath & "ArqFiles\"
    FileName = Dir(FilePath, vbArchive + vbReadOnly + vbHidden)
    Do While FileName <> vbNullString
        ListFiles.AddItem Left(FileName, Len(FileName) - 4)
        FileName = Dir
    Loop
    If ListFiles.ListCount > 0 Then ListFiles.ListIndex = 0

End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set FormPrint = Nothing

End Sub
Private Sub GetPrintingData()

Dim Fs As Object
Dim F As Object
Dim Ts As Object
Dim FileName As String
Dim LineText As String
Dim JustName As String
Dim FolderName As String
Dim Record As Long
Dim Ca As Long
Dim Dpos As Long

    FileName = App.Path
    If Right(FileName, 1) <> "\" Then FileName = FileName & "\"
    FileName = FileName & "ArqFiles\" & ListFiles.List(ListFiles.ListIndex) & ".arq"
    Record = 0
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set F = Fs.GetFile(FileName)
    Set Ts = F.OpenAsTextStream(1, -2)
    Do While Ts.AtEndOfStream <> True
        LineText = Ts.Readline
        Dpos = InStr(LineText, "\")
        If Dpos = 0 Then
            JustName = LineText
            FolderName = "-"
            Else
            For Ca = Len(LineText) To 1 Step -1
                If Mid(LineText, Ca, 1) = "\" Then
                    JustName = Right(LineText, Len(LineText) - Ca)
                    FolderName = Left(LineText, Ca - 1)
                    Exit For
                End If
            Next Ca
        End If
        JustName = JustName & FormatString(Record, 10)
        ListNames.AddItem JustName
        ListFolders.AddItem FolderName
        Record = Record + 1
    Loop
    Ts.Close
    Set Fs = Nothing
    Set F = Nothing
    Set Ts = Nothing

End Sub
Private Sub PrintReport()

On Error GoTo PrintReport_Error
    
Const Twips As Single = 56.7
Dim JustName As String
Dim FolderName As String
Dim Ca As Long
Dim Dpos As Long
Dim OldCap As String
Dim LineNum As Long
Dim Page As Long
Dim Record As Long
Dim OldFontName As String
Dim OldFontSize As Long
Dim OldBold As Boolean
Dim OldItalic As Boolean
Dim OldUnderline As Boolean
Dim OldScaleMode As Long

    OldFontName = Printer.Font.Name
    OldFontSize = Printer.Font.Size
    OldBold = Printer.Font.Bold
    OldItalic = Printer.Font.Italic
    OldUnderline = Printer.Font.Underline
    OldScaleMode = Printer.ScaleMode
    Printer.ScaleMode = vbTwips
    Page = 1
    PrintHeader Page, ListFiles.List(ListFiles.ListIndex) & ".arq"
    LineNum = 0
    For Ca = 0 To ListNames.ListCount - 1
        JustName = ListNames.List(Ca)
        Record = CLng(Right(JustName, 10))
        JustName = Left(JustName, Len(JustName) - 10)
        FolderName = ListFolders.List(Record)
        If UCase(Left(JustName, 1)) <> OldCap Then
            OldCap = UCase(Left(JustName, 1))
            Printer.Font.Bold = True
            Else: Printer.Font.Bold = False
        End If
        Printer.CurrentX = 20 * Twips
        Printer.Print JustName;
        Printer.CurrentX = 90 * Twips
        Printer.Print FolderName
        LineNum = LineNum + 1
        If LineNum > 60 Then
            Printer.Print
            Dpos = Printer.CurrentY
            Printer.DrawWidth = 10
            Printer.Line (20 * Twips, Dpos)-(185 * Twips, Dpos)
            LineNum = 0
            Printer.NewPage
            Page = Page + 1
            PrintHeader Page, ListFiles.List(ListFiles.ListIndex) & ".arq"
        End If
    Next Ca
    Printer.EndDoc
    Printer.Font.Name = OldFontName
    Printer.Font.Size = OldFontSize
    Printer.Font.Bold = OldBold
    Printer.Font.Italic = OldItalic
    Printer.Font.Underline = OldUnderline
    Printer.ScaleMode = OldScaleMode
    
Exit Sub
    
PrintReport_Error:
    MsgBox Err.Description, vbCritical
    Err.Clear

End Sub
Private Function FormatString(Num As Long, Dig As Long) As String

    FormatString = CStr(Num)
    Do While Len(FormatString) < Dig
        FormatString = "0" & FormatString
    Loop

End Function
Sub PrintHeader(PageNum As Long, LDisc As String)

On Error GoTo PrintHeader_Error
    
Const Twips As Single = 56.7
Dim Ypos As Long

    Printer.Font.Name = "Arial"
    Printer.Font.Bold = True
    Printer.Font.Italic = False
    Printer.Font.Underline = False
    Printer.Font.Size = 12
    Printer.CurrentY = 15 * Twips
    Printer.CurrentX = 20 * Twips
    Printer.Print "Disc label:  ";
    Printer.CurrentY = 14.4 * Twips
    Printer.CurrentX = 45 * Twips
    Printer.Font.Size = 14
    Printer.Print LDisc;
    Printer.Font.Size = 12
    Printer.CurrentY = 15 * Twips
    Printer.CurrentX = 170 * Twips
    Printer.Print "Page " & PageNum
    Printer.Print
    Printer.Font.Name = "Arial Narrow"
    Printer.Font.Bold = True
    Printer.Font.Size = 10
    Printer.CurrentX = 20 * Twips
    Printer.Print "File name";
    Printer.CurrentX = 90 * Twips
    Printer.Print "Folder name"
    Printer.Print
    Ypos = Printer.CurrentY - 2.5 * Twips
    Printer.DrawWidth = 10
    Printer.Line (20 * Twips, Ypos)-(185 * Twips, Ypos)
    Printer.Font.Size = 9
    Printer.Font.Bold = False
    Printer.Print
    
Exit Sub
    
PrintHeader_Error:
    MsgBox Err.Description, vbCritical
    Err.Clear

End Sub
