VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildStaMultiAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add List Detail"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2880
      TabIndex        =   7
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   4320
      TabIndex        =   6
      Top             =   4080
      Width           =   1305
   End
   Begin VB.ComboBox cboInUse 
      Height          =   315
      ItemData        =   "stcmuadd.frx":0000
      Left            =   1920
      List            =   "stcmuadd.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtUserDef1 
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox txtUserDef2 
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox txtSeqNum 
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   16
      Top             =   4545
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4921
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "03/12/01"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "16:21"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInUse 
      Alignment       =   1  'Right Justify
      Caption         =   "In Use :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblUserDef2 
      Alignment       =   1  'Right Justify
      Caption         =   "User Defined 2 :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblUserDef1 
      Alignment       =   1  'Right Justify
      Caption         =   "User Defined 1 :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Sequence Number :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Description :"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Code :"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "List :"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblList 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmChildStaMultiAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrListName As String
Dim mstrCode As String
Dim mstrDescription As String
Dim mstrUserDef1 As String
Dim mstrUserDef2 As String
Dim mbooInUse As Boolean
Dim mlngSeqNum As Long
Dim mstrTranRoute As String
Public Property Let TranRoute(pstrTranRoute As String)
    mstrTranRoute = pstrTranRoute
End Property
Public Property Get TranRoute() As String
    TranRoute = mstrTranRoute
End Property
Public Property Let ListName(pstrListName As String)
    mstrListName = (pstrListName)
End Property
Public Property Get ListName() As String
    ListName = mstrListName
End Property
Public Property Let Code(pstrCode As String)
    mstrCode = (pstrCode)
End Property
Public Property Get Code() As String
    Code = mstrCode
End Property
Public Property Let Description(pstrDescription As String)
    mstrDescription = (pstrDescription)
End Property
Public Property Get Description() As String
    Description = mstrDescription
End Property
Public Property Let SeqNum(plngSeqNum As Long)
    mlngSeqNum = (plngSeqNum)
End Property
Public Property Get SeqNum() As Long
    SeqNum = mlngSeqNum
End Property
Public Property Let UserDef1(pstrUserDef1 As String)
    mstrUserDef1 = (pstrUserDef1)
End Property
Public Property Get UserDef1() As String
    UserDef1 = mstrUserDef1
End Property
Public Property Let UserDef2(pstrUserDef2 As String)
    mstrUserDef2 = (pstrUserDef2)
End Property
Public Property Get UserDef2() As String
    UserDef2 = mstrUserDef2
End Property
Public Property Let InUse(pstrInUse As Boolean)
    mbooInUse = (pstrInUse)
End Property
Public Property Get InUse() As Boolean
    InUse = mbooInUse
End Property
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()
Dim lstrSQL As String

    On Error GoTo ErrHandler
    
    If Trim$(txtCode) = "" Or Trim$(txtDescription) = "" Then
        MsgBox "You must enter a Code vlaue and a Description Value!", , gconstrTitlPrefix & "Mandatory Field"
        Exit Sub
    End If
    
    Select Case mstrTranRoute
    Case gconstrAdminRoute & "EDIT"
        lstrSQL = "UPDATE ListsMaster INNER JOIN ListDetailsMaster ON ListsMaster.ListNum" & _
            "= ListDetailsMaster.ListNum SET " & _
            "ListDetailsMaster.Description = '" & OneSpace(txtDescription) & _
            "', ListDetailsMaster.SequenceNum = " & Val(txtSeqNum) & _
            ", ListDetailsMaster.UserDef1 = '" & OneSpace(txtUserDef1) & _
            "', ListDetailsMaster.UserDef2 = '" & OneSpace(txtUserDef2) & _
            "', ListDetailsMaster.InUse = " & cboInUse & " " & _
            "WHERE (((ListsMaster.ListName)='" & lblList & _
            "') AND ((ListDetailsMaster.ListCode)='" & txtCode & "'));"
        gdatCentralDatabase.Execute lstrSQL
    Case gconstrAdminRoute & "ADD"
        lstrSQL = "INSERT INTO ListDetailsMaster ( ListNum, ListCode, SequenceNum, " & _
            "Description, UserDef1, UserDef2, InUse ) SELECT ListsMaster.ListNum, " & _
            "'" & Trim$(txtCode) & "' AS Expr1, " & Val(txtSeqNum) & " AS Expr2, '" & OneSpace(txtDescription) & _
            "' AS Expr3, '" & OneSpace(txtUserDef1) & "' AS Expr4, '" & OneSpace(txtUserDef2) & _
            "' AS Expr5, '" & cboInUse & "' AS Expr6 From ListsMaster " & _
            "WHERE (((ListsMaster.ListName)='" & lblList & "'));"
        gdatCentralDatabase.Execute lstrSQL
    End Select
    
    mstrCode = txtCode
    mstrDescription = txtDescription
    mlngSeqNum = Val(txtSeqNum)
    
    Unload Me
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmChilStaMultiAdd.cmdOK_Click", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    lblList = mstrListName
    
    Select Case mstrTranRoute
    Case gconstrConfigRoute & "ADD"
        Me.Caption = "Add List Detail"
        cboInUse.Visible = False
        lblInUse.Visible = False
    Case gconstrAdminRoute & "ADD"
        Me.Caption = "Add List Detail"
    Case gconstrAdminRoute & "EDIT"
        Me.Caption = "Edit List Detail"
        txtCode = mstrCode
        txtCode.Enabled = False
        'Populate Fields from DB
        GetListDetailItem mstrListName, mstrCode
    End Select
    
End Sub

Private Sub txtCode_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtDescription_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtSeqNum_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtSeqNum_LostFocus()

    If Not IsNumeric(txtSeqNum) Then txtSeqNum = 0
    
End Sub

Private Sub txtUserDef1_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtUserDef2_GotFocus()

    SetSelected Me
    
End Sub
Sub GetListDetailItem(pstrListName As String, pstrListDetailCode As String)
Dim lstrSQL As String
Dim lsnaLists As Recordset

    On Error GoTo ErrHandler
    

    lstrSQL = "SELECT ListsMaster.ListName, ListDetailsMaster.ListCode, " & _
        "ListDetailsMaster.Description, ListDetailsMaster.UserDef1, " & _
        "ListDetailsMaster.UserDef2, ListDetailsMaster.InUse, " & _
        "ListDetailsMaster.SequenceNum FROM ListsMaster INNER JOIN " & _
        "ListDetailsMaster ON ListsMaster.ListNum = ListDetailsMaster.ListNum " & _
        "Where (((ListsMaster.ListName) = '" & Trim$(pstrListName) & _
        "' AND (ListDetailsMaster.ListCode)='" & Trim$(pstrListDetailCode) & "'));"

    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    If Not lsnaLists.EOF Then
        txtDescription = lsnaLists.Fields("Description")
        txtSeqNum = lsnaLists.Fields("SequenceNum")
        txtUserDef1 = lsnaLists.Fields("UserDef1") & ""
        txtUserDef2 = lsnaLists.Fields("UserDef2") & ""
        cboInUse = CStr(CBool(lsnaLists.Fields("InUse")))
        lsnaLists.MoveNext
    End If
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetListDetailItem", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub



