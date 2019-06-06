VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1305
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   360
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1305
   End
   Begin VB.ListBox lstList 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   2715
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
            TextSave        =   "18/07/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "15:52"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChildOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrCode As String
Dim mstrList As String
Dim lstrListDesc() As String
Dim lstrScreenHelpFile As String
Public Property Let Code(pstrCode As String)

    mstrCode = pstrCode

End Property
Public Property Get Code() As String

    Code = mstrCode

End Property
Public Property Let List(pstrList As String)

    mstrList = pstrList

End Property
Public Property Get List() As String

    List = mstrList

End Property

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdSelect_Click()
    
    mstrCode = Trim$(NotNull(lstList, lstrListDesc))
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me

    FillList mstrList, lstList, lstrListDesc()
    SelectListItem Trim$(mstrCode), lstList, lstrListDesc()
    
    SetupHelpFileReqs
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/ChildGenericOptions.xml>WhatsScreen"
        
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdSelect.WhatsThisHelpID = IDH_STANDARD_NEXT
    lstList.WhatsThisHelpID = IDH_CHIGENOPS_LIST
    
End Sub
