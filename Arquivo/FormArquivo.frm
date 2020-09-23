VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FormArquivo 
   Caption         =   "Arquivo"
   ClientHeight    =   3555
   ClientLeft      =   1350
   ClientTop       =   2655
   ClientWidth     =   5070
   Icon            =   "FormArquivo.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "FormArquivo.frx":0CCA
   ScaleHeight     =   3555
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ListIcons 
      Left            =   4395
      Top             =   2415
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormArquivo.frx":9595
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormArquivo.frx":9E6F
            Key             =   "Catalog"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormArquivo.frx":A749
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   1535
      ButtonWidth     =   1349
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ListIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
            Key             =   "Find"
            Object.ToolTipText     =   "Find File/Folder"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Catalog"
            Key             =   "Catalog"
            Object.ToolTipText     =   "Catalog Disc Contents"
            ImageKey        =   "Catalog"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Print Disc Contents"
            ImageKey        =   "Print"
         EndProperty
      EndProperty
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   1000
         Left            =   3900
         TabIndex        =   1
         Top             =   -60
         Width           =   1065
         ExtentX         =   1879
         ExtentY         =   1764
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileFind 
         Caption         =   "&Find File/Folder..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileCatalog 
         Caption         =   "&Catalog Disc Contents..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print Disc Contents..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "FormArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()

Dim Anim As String
Anim = App.Path
If Right(Anim, 1) <> "\" Then Anim = Anim & "\"
Anim = Anim & "ArqAnims.gif"
WebBrowser1.Navigate "about:<html><body scroll=NO><body bgcolor=""#FFFFFF""><BODY TOPMARGIN=""0"" LEFTMARGIN=""0"" MARGINWIDTH=""0"" MARGINHEIGHT=""0""><IMG src=""" & Anim & """></body></html></p>"


End Sub

Private Sub mnuAbout_Click()

SendKeys "{F1}"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Find"
            mnuFileFind_Click
        Case "Catalog"
            mnuFileCatalog_Click
        Case "Print"
            mnuFilePrint_Click
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub
Private Sub mnuFilePrint_Click()

    FormPrint.Show vbModal, Me

End Sub
Private Sub mnuFileCatalog_Click()

    FormCatalog.Show vbModal, Me

End Sub
Private Sub mnuFileFind_Click()

    FormFind.Show vbModal, Me

End Sub
