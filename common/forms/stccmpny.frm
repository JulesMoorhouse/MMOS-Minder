VERSION 5.00
Begin VB.Form frmStaticCompany 
   Caption         =   "Specify values"
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
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   25
      Top             =   6317
      Width           =   4095
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   23
      Top             =   7110
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1244
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   20
      Top             =   5586
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   17
      Top             =   4855
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   14
      Top             =   4124
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   11
      Top             =   3393
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   2662
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   1931
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   7
      Left            =   600
      TabIndex        =   27
      Top             =   6317
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   26
      Top             =   6602
      Width           =   5895
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   19
      Top             =   5871
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   6
      Left            =   600
      TabIndex        =   18
      Top             =   5586
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   16
      Top             =   5140
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   15
      Top             =   4855
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   13
      Top             =   4409
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Top             =   4124
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   3678
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   3393
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   2947
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   2662
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   2216
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   1931
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   1485
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmStaticCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lintStartIndex As Integer
Dim lintEndIndex As Integer

Dim mfrmCallingForm As Object
Dim mstrRoute As String
Dim mfrmFinalForm As Object
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
Dim lintArrInc As Integer
    
    Select Case Route
    Case gconstrAdminRoute
        gstrButtonRoute = gconstrReferenceData
        UnloadLastForm
        Set gstrCurrentLoadedForm = mfrmCallingForm
        mdiMain.DrawButtonSet gstrButtonRoute
        mfrmCallingForm.Show

    Case gconstrConfigRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            UpdateObjects lintArrInc + lintStartIndex, txtDescription(lintArrInc)
        Next lintArrInc
        
        Unload Me
        mfrmCallingForm.Show

    End Select

End Sub
Private Sub cmdProceed_Click()
Dim lintArrInc As Integer
    
    For lintArrInc = 0 To lintEndIndex - lintStartIndex
        If CheckObjects(lintArrInc + lintStartIndex, txtDescription(lintArrInc)) = False Then
            Exit Sub
        End If
    Next lintArrInc
    
    Me.Enabled = False
    
    Select Case Route
    Case gconstrAdminRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            If gstrSystemLists(lintArrInc).strDescValue <> txtDescription(lintArrInc) Then
                UpdateListDetailWithObject lintArrInc + lintStartIndex, txtDescription(lintArrInc)
            End If
        Next lintArrInc
        
        gstrButtonRoute = gconstrReferenceData
        UnloadLastForm
        Set gstrCurrentLoadedForm = mfrmCallingForm
        mdiMain.DrawButtonSet gstrButtonRoute
        Me.Enabled = True
        mfrmCallingForm.Show

    Case gconstrConfigRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            UpdateObjects lintArrInc + lintStartIndex, txtDescription(lintArrInc)
        Next lintArrInc
        
        Unload Me
        
        frmStaticPForce.FinalForm = mfrmFinalForm
        frmStaticPForce.CallingForm = Me
        frmStaticPForce.Route = gconstrConfigRoute
        Me.Enabled = True
        frmStaticPForce.Show

    End Select
    
End Sub

Private Sub Form_Load()
Dim lintArrInc As Integer
    
    lintStartIndex = gconintStaticCompanyName
    lintEndIndex = gconintStaticCompanyTelNum
    
    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    Select Case mstrRoute
    Case gconstrAdminRoute
        cmdProceed.Caption = "&Save"
        cmdBack.Caption = "&Cancel"
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            FillObjectWithListValue lintArrInc + lintStartIndex, lblTopic(lintArrInc), txtDescription(lintArrInc)
            lblExampleDesc(lintArrInc) = ""
        Next lintArrInc

    Case gconstrConfigRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            PopulateObjects lintArrInc + lintStartIndex, lblTopic(lintArrInc), txtDescription(lintArrInc), lblExampleDesc(lintArrInc)
        Next lintArrInc
        
    End Select
    
    ShowBanner Me, mstrRoute
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim llngSpaceAdj As Long
Dim lintArrInc As Integer
Const lconTopBoxPos = 1200
Dim llngAvailSpace As Long
Dim llngLastTop As Long

    llngSpaceAdj = 2600
    llngAvailSpace = (Me.Height - llngSpaceAdj) - ((txtDescription(0).Height + lblExampleDesc(0).Height + 75) * 8)
    
    With cmdProceed
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdProceed.Left - (cmdBack.Width + 120)
    End With
    llngLastTop = lconTopBoxPos
    
    For lintArrInc = 0 To lintEndIndex - lintStartIndex
        txtDescription(lintArrInc).Top = llngLastTop
        lblTopic(lintArrInc).Top = llngLastTop
        lblExampleDesc(lintArrInc).Top = txtDescription(lintArrInc).Height + llngLastTop
        llngLastTop = lblExampleDesc(lintArrInc).Top + lblExampleDesc(lintArrInc).Height + 75
        llngLastTop = llngLastTop + (llngAvailSpace / 8)
    Next lintArrInc
    
End Sub

