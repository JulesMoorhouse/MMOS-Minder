VERSION 5.00
Begin VB.Form frmChildAbortOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Abort Options"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save && Set to Cancelled"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort This Order"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmChildAbortOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrAbortOption As String
Dim mlngStyle As Long
Public Property Get Style() As Long

    Style = mlngStyle
    
End Property
Public Property Let Style(plngStyle As Long)

    mlngStyle = plngStyle

End Property
Public Property Get AbortOption() As String

    AbortOption = mstrAbortOption
    
End Property
Public Property Let AbortOption(pstrAbortOption As String)

    mstrAbortOption = pstrAbortOption

End Property

Private Sub cmdAbort_Click()

    mstrAbortOption = "ABORT"
    Unload Me
    
End Sub

Private Sub cmdBack_Click()

    mstrAbortOption = "BACK"
    Unload Me
    
End Sub

Private Sub cmdSave_Click()

    mstrAbortOption = "SAVE"
    Unload Me
    
End Sub

Private Sub Form_Load()

    Select Case mlngStyle
    Case 1
        cmdSave.Enabled = False
    Case 0
        'All
        cmdSave.Enabled = True
    End Select
    
End Sub
