VERSION 5.00
Begin VB.Form frmStaMultiMarket 
   Caption         =   "Add values"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   10545
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLists 
      BorderStyle     =   0  'None
      Height          =   2775
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   8415
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   360
         Index           =   1
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1200
         Width           =   1305
      End
      Begin VB.ListBox lstCode 
         Height          =   1815
         Index           =   1
         ItemData        =   "stcmumark.frx":0000
         Left            =   120
         List            =   "stcmumark.frx":0007
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.ListBox lstDescription 
         Height          =   1815
         Index           =   1
         Left            =   1560
         TabIndex        =   20
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   360
         Index           =   1
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   1305
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   360
         Index           =   1
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   1305
      End
      Begin VB.ListBox lstSeqNum 
         Height          =   1815
         Index           =   1
         Left            =   5640
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblCode 
         Caption         =   "Code"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblListName 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label lblSeqNum 
         Caption         =   "Sequence Num"
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblItemCount 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.Frame fraLists 
      BorderStyle     =   0  'None
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   8415
      Begin VB.ListBox lstSeqNum 
         Height          =   1815
         Index           =   0
         Left            =   5640
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   360
         Index           =   0
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   1305
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   360
         Index           =   0
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1305
      End
      Begin VB.ListBox lstDescription 
         Height          =   1815
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   4095
      End
      Begin VB.ListBox lstCode 
         Height          =   1815
         Index           =   0
         ItemData        =   "stcmumark.frx":0017
         Left            =   120
         List            =   "stcmumark.frx":001E
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   360
         Index           =   0
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label lblWarning 
         Caption         =   "Remember, values may NEVER be deleted, just set to ""Not In Use""."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label lblItemCount 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblSeqNum 
         Caption         =   "Sequence Num"
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblListName 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblCode 
         Caption         =   "Code"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   2
      Top             =   7110
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1244
   End
End
Attribute VB_Name = "frmStaMultiMarket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const lintStartIndex = 4
Const lintEndIndex = 5
Dim mfrmCallingForm As Object
Dim mstrRoute As String
Dim mfrmFinalForm As Object
Dim mfrmNextForm As Object
Public Property Let NextForm(pstrNextForm As Object)

    Set mfrmNextForm = pstrNextForm

End Property
Public Property Get NextForm() As Object

    NextForm = mfrmNextForm
    
End Property
Public Property Let FinalForm(pstrFinalForm As Object)

    Set mfrmFinalForm = pstrFinalForm

End Property
Public Property Get FinalForm() As Object

    FinalForm = mfrmFinalForm
    
End Property
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property
Public Property Let CallingForm(pstrCallingForm As Object)

    Set mfrmCallingForm = pstrCallingForm

End Property
Public Property Get CallingForm() As Object

    CallingForm = mfrmCallingForm
    
End Property
Private Sub cmdBack_Click()

    Select Case Route
    Case gconstrAdminRoute
        gstrButtonRoute = gconstrReferenceData
        UnloadLastForm
        Set gstrCurrentLoadedForm = mfrmCallingForm
        mdiMain.DrawButtonSet gstrButtonRoute
        mfrmCallingForm.Show
    Case gconstrConfigRoute
        Unload Me
        mfrmCallingForm.Show
    End Select

End Sub
Sub SetListIndex(pintListIndex As Integer, pintControlIndex As Integer)

    lstCode(pintControlIndex).ListIndex = pintListIndex
    lstDescription(pintControlIndex).ListIndex = pintListIndex
    lstSeqNum(pintControlIndex).ListIndex = pintListIndex
    
End Sub
Private Sub Label2_Click()

End Sub

Private Sub cmdAdd_Click(Index As Integer)

    If lstCode(Index).ListCount = 10 And mstrRoute = gconstrConfigRoute Then
        MsgBox "You may only enter 10 items, in this setup phase!" & vbCrLf & "You may add more later!", , gconstrTitlPrefix & "Item Check"
        Exit Sub
    End If
    
    With frmChildStaMultiAdd
        .TranRoute = mstrRoute & "ADD"
        .ListName = lblListName(Index)
        .Code = ""
        .Description = ""
        .SeqNum = 0
        .UserDef1 = ""
        .UserDef2 = ""
        .InUse = True
        frmChildStaMultiAdd.Show vbModal
        If mstrRoute = gconstrConfigRoute Then
            lstCode(Index).AddItem .Code
            lstDescription(Index).AddItem .Description
            lstSeqNum(Index).AddItem .SeqNum
        ElseIf mstrRoute = gconstrAdminRoute Then
            RefreshListboxes
        End If
    End With
    
    lblItemCount(Index) = "(" & lstCode(Index).ListCount & ") items."
    
End Sub

Private Sub xcmdBack_Click()

End Sub

Private Sub cmdEdit_Click(Index As Integer)

    With frmChildStaMultiAdd
        'Route will only ever be Admin Screen edit
        .TranRoute = mstrRoute & "EDIT"
        .ListName = lblListName(Index)
        .Code = lstCode(Index)
        
        frmChildStaMultiAdd.Show vbModal
        
        RefreshListboxes
    End With
    
    lblItemCount(Index) = "(" & lstCode(Index).ListCount & ") items."
    
End Sub
Sub RefreshListboxes()
Dim lintArrInc As Integer

    For lintArrInc = 0 To lintEndIndex - lintStartIndex
        FillMultiObjectWithListValues lintArrInc + lintStartIndex, lstCode(lintArrInc), _
            lstDescription(lintArrInc), lstSeqNum(lintArrInc), lblListName(lintArrInc)
        lblItemCount(lintArrInc) = "(" & lstCode(lintArrInc).ListCount & ") items."
    Next lintArrInc
        
End Sub
Public Sub cmdProceed_Click()
Dim lintArrInc As Integer

    For lintArrInc = 0 To lintEndIndex - lintStartIndex
        If lstCode(lintArrInc).ListCount = 0 Then
            MsgBox "You must have at least one value in each list!", , gconstrTitlPrefix & "List Item Check"
            Exit Sub
        End If
    Next lintArrInc
    
    Me.Enabled = False
    
    Select Case Route
    Case gconstrAdminRoute
        gstrButtonRoute = gconstrReferenceData
        UnloadLastForm
        Set gstrCurrentLoadedForm = mfrmCallingForm
        mdiMain.DrawButtonSet gstrButtonRoute
        Me.Enabled = True
        mfrmCallingForm.Show
    Case gconstrConfigRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            ReStackArray lintArrInc + lintStartIndex, lstCode(lintArrInc), _
                    lstDescription(lintArrInc), lstSeqNum(lintArrInc)
        Next lintArrInc
        
        Unload Me
        mfrmNextForm.FinalForm = mfrmFinalForm
        mfrmNextForm.CallingForm = Me
        mfrmNextForm.Route = gconstrConfigRoute
        Me.Enabled = True
        mfrmNextForm.Show
    End Select
    
End Sub

Private Sub cmdRemove_Click(Index As Integer)
Dim lintRetVal

    If lstCode(Index).ListIndex = -1 Then
        MsgBox "You must select an item to remove first!", , gconstrTitlPrefix & "No Item selected!"
        Exit Sub
        
    End If
    
    lintRetVal = MsgBox("Do you wish to remove this item?", vbYesNo, gconstrTitlPrefix & "Item Selection")
    If lintRetVal = vbYes Then
        lstCode(Index).RemoveItem lstCode(Index).ListIndex
        lstDescription(Index).RemoveItem lstDescription(Index).ListIndex
        lstSeqNum(Index).RemoveItem lstSeqNum(Index).ListIndex
        lblItemCount(Index) = "(" & lstCode(Index).ListCount & ") items."
    End If
        
End Sub

Private Sub Form_Load()
Dim lintArrInc As Integer

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me

    Select Case mstrRoute
    Case gconstrAdminRoute
        lblWarning.Visible = True
        cmdBack.Visible = False
        cmdProceed.Caption = "&Back"
        
        RefreshListboxes
        
        cmdRemove(0).Visible = False
        cmdRemove(1).Visible = False
    Case gconstrConfigRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            PopulateMultiObjects lintArrInc + lintStartIndex, lstCode(lintArrInc), _
                lstDescription(lintArrInc), lstSeqNum(lintArrInc), lblListName(lintArrInc)
            lblItemCount(lintArrInc) = "(" & lstCode(lintArrInc).ListCount & ") items."
        Next lintArrInc
        cmdEdit(0).Visible = False
        cmdRemove(0).Top = cmdEdit(0).Top
        cmdEdit(1).Visible = False
        cmdRemove(1).Top = cmdEdit(1).Top
    End Select
    
    ShowBanner Me, mstrRoute
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub lstCode_Click(Index As Integer)

    SetListIndex lstCode(Index).ListIndex, Index
    
End Sub

Private Sub lstCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    SetListIndex lstCode(Index).ListIndex, Index

End Sub

Private Sub lstDescription_Click(Index As Integer)

    SetListIndex lstDescription(Index).ListIndex, Index

End Sub

Private Sub lstDescription_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    SetListIndex lstDescription(Index).ListIndex, Index
    
End Sub

Private Sub lstSeqNum_Click(Index As Integer)

    SetListIndex lstSeqNum(Index).ListIndex, Index
    
End Sub

Private Sub lstSeqNum_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    SetListIndex lstSeqNum(Index).ListIndex, Index
    
End Sub


Private Sub lstCode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetListIndex lstCode(Index).ListIndex, Index
    
End Sub

Private Sub lstDescription_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetListIndex lstDescription(Index).ListIndex, Index
    
End Sub

Private Sub lstSeqNum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetListIndex lstSeqNum(Index).ListIndex, Index
    
End Sub

Private Sub Form_Resize()
Dim llngSpaceAdj As Long
Dim llng100Percent As Long
Const lconCodePercent = 21.5
Const lconDescPercent = 60.5
Const lconSeqPercent = 18.5

    llngSpaceAdj = 2600
    With fraLists(0)
        .Height = (Me.Height - llngSpaceAdj) / 2
        .Width = Me.Width - 240
    End With

    With fraLists(1)
        .Top = fraLists(0).Top + fraLists(0).Height + 225
        .Height = fraLists(0).Height
        .Width = Me.Width - 240
    End With

    cmdAdd(0).Left = (fraLists(0).Width - cmdAdd(0).Width) - 120
    cmdEdit(0).Left = cmdAdd(0).Left
    cmdRemove(0).Left = cmdAdd(0).Left
    
    cmdAdd(1).Left = (fraLists(1).Width - cmdAdd(1).Width) - 120
    cmdEdit(1).Left = cmdAdd(1).Left
    cmdRemove(1).Left = cmdAdd(1).Left
    
    lblItemCount(0).Top = fraLists(0).Height - lblItemCount(0).Height
    lblItemCount(1).Top = fraLists(1).Height - lblItemCount(1).Height
    
    lstCode(0).Height = lblItemCount(0).Top - lstCode(0).Top
    lstCode(1).Height = lblItemCount(1).Top - lstCode(1).Top
    lstDescription(0).Height = lstCode(0).Height
    lstDescription(1).Height = lstCode(1).Height
    lstSeqNum(0).Height = lstCode(0).Height
    lstSeqNum(1).Height = lstCode(1).Height
    
    llng100Percent = (cmdAdd(0).Left - (195 + 120)) / 100
    lstCode(0).Width = llng100Percent * lconCodePercent
    lstDescription(0).Left = (lstCode(0).Width + lstCode(0).Left) - 15
    lstDescription(0).Width = llng100Percent * lconDescPercent
    lstSeqNum(0).Left = (lstDescription(0).Width + lstDescription(0).Left) - 15
    lstSeqNum(0).Width = llng100Percent * lconSeqPercent
    lblDescription(0).Left = lstDescription(0).Left
    lblSeqNum(0).Left = lstSeqNum(0).Left
    
    llng100Percent = (cmdAdd(1).Left - (195 + 120)) / 100
    lstCode(1).Width = llng100Percent * lconCodePercent
    lstDescription(1).Left = (lstCode(1).Width + lstCode(1).Left) - 15
    lstDescription(1).Width = llng100Percent * lconDescPercent
    lstSeqNum(1).Left = (lstDescription(1).Width + lstDescription(1).Left) - 15
    lstSeqNum(1).Width = llng100Percent * lconSeqPercent
    lblDescription(1).Left = lstDescription(1).Left
    lblSeqNum(1).Left = lstSeqNum(1).Left
    
    With cmdProceed
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdProceed.Left - (cmdBack.Width + 120)
    End With
    
End Sub
