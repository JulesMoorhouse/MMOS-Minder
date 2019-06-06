VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildPrinter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Printer options"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   360
      Left            =   3360
      TabIndex        =   4
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "This option may not be relevant!"
      Height          =   1455
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
      Begin VB.OptionButton optLinesAPage 
         Caption         =   "I have 60 lines per page (A4 Deskjet)"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Tag             =   "60"
         Top             =   960
         Width           =   3135
      End
      Begin VB.OptionButton optLinesAPage 
         Caption         =   "I have 77 lines per page (A4)"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Tag             =   "77"
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optLinesAPage 
         Caption         =   "I have 66 lines per page"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "66"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.ListBox lstLPT 
      Height          =   840
      ItemData        =   "Printer.frx":0000
      Left            =   3000
      List            =   "Printer.frx":000D
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3069
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "25/03/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "16:38"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Please select a printer port:-"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmChildPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintLinesPerPage As Integer
Dim mstrLPTPort As String
Public Property Let LinesPerPage(pintLinesPerpage As Integer)

    mintLinesPerPage = pintLinesPerpage

End Property
Public Property Get LinesPerPage() As Integer

    LinesPerPage = mintLinesPerPage

End Property
Public Property Let LPTPort(pstrLPTPort As String)

    mstrLPTPort = pstrLPTPort

End Property
Public Property Get LPTPort() As String

    LPTPort = mstrLPTPort

End Property

Private Sub cmdCancel_Click()
    
End Sub

Private Sub cmdHelp_Click()

    RunNDontWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1014.htm"

End Sub

Private Sub cmdPrint_Click()
Dim lintArrInc As Integer

    For lintArrInc = 0 To optLinesAPage.Count - 1
        If optLinesAPage.Item(lintArrInc).Value = True Then
            mintLinesPerPage = Val(optLinesAPage(lintArrInc).Tag)
        End If
    Next lintArrInc

    For lintArrInc = 0 To lstLPT.ListCount - 1
        If lstLPT.Selected(lintArrInc) = True Then
            mstrLPTPort = lstLPT.List(lintArrInc)
        End If
    Next lintArrInc

    Unload Me
End Sub

Private Sub Form_Load()
Dim lintArrInc As Integer

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    Select Case mintLinesPerPage
    Case 66
        optLinesAPage.Item(0).Value = True
    Case 77
        optLinesAPage.Item(1).Value = True
    End Select

    For lintArrInc = 0 To lstLPT.ListCount
        If lstLPT.List(lintArrInc) = UCase$(mstrLPTPort) Then
            lstLPT.ListIndex = lintArrInc
        End If
    Next lintArrInc
    
End Sub
