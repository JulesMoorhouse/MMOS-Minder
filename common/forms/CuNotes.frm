VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildCuNotes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customer Notepad"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8475
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox txtNotes 
      Height          =   1935
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   8175
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   4920
      Top             =   2520
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1305
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   2955
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9763
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "16/04/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "13:03"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   $"CuNotes.frx":0000
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmChildCuNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbooFoundNotes As Boolean
Dim lstrScreenHelpFile As String
Private Sub cmdBack_Click()

End Sub

Private Sub cmdClose_Click()

    If lbooFoundNotes = False Then
        AddNewCustomerNote gstrCustomerAccount.lngCustNum, txtNotes
    Else
        UpdateCustomerNotes gstrCustomerAccount.lngCustNum, txtNotes
    End If
    
    Unload Me

End Sub

Private Sub cmdHelp_Click()

    If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
        RunNDontWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1008.htm"
    Else
        glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    End If
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub
Private Sub Form_Load()
Dim lstrNotes As String
    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    lbooFoundNotes = GetCustomerNote(gstrCustomerAccount.lngCustNum, lstrNotes)
    txtNotes = lstrNotes
    
    SetupHelpFileReqs
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
End Sub

Private Sub txtNotes_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub

Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/ChildNotes.xml>WhatsScreen"
        
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdClose.WhatsThisHelpID = IDH_STANDARD_BACK
    txtNotes.WhatsThisHelpID = IDH_STANDARD_CUNOTES
    
End Sub
