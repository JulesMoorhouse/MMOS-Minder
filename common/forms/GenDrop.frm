VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildGenericDropdown 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GenericDropdown"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   1305
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   1305
   End
   Begin VB.ComboBox cboGenericList 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   2130
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3043
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "07/03/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:10"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblGenericLabelCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblGenericLabel 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "frmChildGenericDropdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSQL As String
Dim mstrLabel As String
Dim mstrDescField As String
Dim mstrCodeField As String
Dim mstrReturnCode As String
Dim lstrCode() As String
Dim mstrLabelCaption As String
Dim mbooCancelled As Boolean
Dim mstrFormCaption As String
Dim mbooAddStar As Boolean
Dim mstrDB As String
Public Property Get DB() As String

    DB = mstrDB
    
End Property
Public Property Let DB(pstrDB As String)

    mstrDB = pstrDB

End Property
Public Property Get AddStar() As Boolean

    AddStar = mbooAddStar
    
End Property
Public Property Let AddStar(pstrAddStar As Boolean)

    mbooAddStar = pstrAddStar

End Property
Public Property Get FormCaption() As String

    FormCaption = mstrFormCaption
    
End Property
Public Property Let FormCaption(pstrFormCaption As String)

    mstrFormCaption = pstrFormCaption

End Property
Public Property Let Cancelled(pstrCancelled As Boolean)

    mbooCancelled = pstrCancelled

End Property
Public Property Get Cancelled() As Boolean

    Cancelled = mbooCancelled
    
End Property
Public Property Get LabelCaption() As String

    LabelCaption = mstrLabelCaption
    
End Property
Public Property Let LabelCaption(pstrLabelCaption As String)

    mstrLabelCaption = pstrLabelCaption

End Property
Public Property Let ReturnCode(pstrLabel As String)

    mstrReturnCode = pstrLabel

End Property
Public Property Get ReturnCode() As String

    ReturnCode = mstrReturnCode
    
End Property
Public Property Let CodeField(pstrLabel As String)

    mstrCodeField = pstrLabel

End Property
Public Property Get CodeField() As String

    CodeField = mstrCodeField
    
End Property
Public Property Let DescField(pstrLabel As String)

    mstrDescField = pstrLabel

End Property
Public Property Get DescField() As String

    DescField = mstrDescField
    
End Property
Public Property Let LabelStr(pstrLabel As String)

    mstrLabel = pstrLabel

End Property
Public Property Get LabelStr() As String

    LabelStr = mstrLabel
    
End Property
Public Property Let SQL(pstrSQL As String)

    mstrSQL = pstrSQL

End Property
Public Property Get SQL() As String

    SQL = mstrSQL
    
End Property

Private Sub cmdCancel_Click()

    mbooCancelled = True
    Unload Me
    
End Sub

Private Sub cmdSelect_Click()

    mbooCancelled = False
    mstrReturnCode = Trim$(NotNull(cboGenericList, lstrCode))
    Unload Me
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    mbooCancelled = True
    Me.Caption = mstrFormCaption
    lblGenericLabel = mstrLabel
    lblGenericLabelCaption = mstrLabelCaption
    FillGenericList cboGenericList, lstrCode, mstrSQL, mstrCodeField, mstrDescField, mbooAddStar, mstrDB
        
    cboGenericList.ListIndex = 0
    
End Sub
