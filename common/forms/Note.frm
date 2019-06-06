VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildNote 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Notes"
   ClientHeight    =   2760
   ClientLeft      =   30
   ClientTop       =   225
   ClientWidth     =   7440
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   2040
      Width           =   375
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   4320
      Top             =   2040
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1305
   End
   Begin VB.TextBox txtNote 
      Height          =   1335
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   7212
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   2490
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7938
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "14/08/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "14:08"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNoteType 
      Caption         =   "Note Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7212
   End
End
Attribute VB_Name = "frmChildNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrNoteText As String
Dim mstrNoteType As String
Dim lstrScreenHelpFile As String
Public Property Let NoteText(pstrNote As String)

    mstrNoteText = Trim$(pstrNote)

End Property
Public Property Get NoteText() As String

    NoteText = Trim$(mstrNoteText)

End Property
Public Property Let NoteType(pstrNoteType As String)

    mstrNoteType = pstrNoteType
    
End Property
Public Property Get NoteType() As String

    NoteType = ProperCase(Trim$(mstrNoteType))

End Property

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub
Private Sub cmdOK_Click()

    mstrNoteText = txtNote.Text
    Unload Me
End Sub

Private Sub Form_Activate()

    If mstrNoteType = "Consignment Note Comments" Then
        If gstrAdviceNoteOrder.intNumOfParcels > 0 Then
            MsgBox "This order has already been packed and therefore the consignment note may not be altered!", vbInformation, gconstrTitlPrefix & "Consignment Note"
            txtNote.Locked = True
            txtNote.Enabled = False
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
        
    txtNote.Text = mstrNoteText
    lblNoteType.Caption = mstrNoteType
    
    SetupHelpFileReqs
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub

Private Sub txtNote_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)

End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/ChildNotes.xml>WhatsScreen"
        
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdOK.WhatsThisHelpID = IDH_STANDARD_BACK
    txtNote.WhatsThisHelpID = IDH_STANDARD_CONSNOTE
    
End Sub
