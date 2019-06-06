VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPackaging 
   Caption         =   "Please Select an Order..."
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10515
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.CommandButton cmdViewAdviceNote 
      Caption         =   "&View Advice Note"
      Height          =   360
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1200
      Width           =   1543
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   3720
      Top             =   6960
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7235
      Width           =   1332
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&Confirm"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7235
      Width           =   1332
   End
   Begin VB.Data datOrderLineMaster 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid dbgOrderLinesMaster 
      Bindings        =   "Packaging.frx":0000
      Height          =   3330
      Left            =   120
      OleObjectBlob   =   "Packaging.frx":0021
      TabIndex        =   2
      Top             =   3240
      Width           =   10245
   End
   Begin VB.TextBox txtSearchCriteria 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2412
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1332
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7235
      Width           =   1332
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   30
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin VB.Label lblOrderStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   34
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Order Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblFoundNumber 
      Caption         =   "Found 0 records"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Courier:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Deliver To :-"
      Height          =   255
      Left            =   7680
      TabIndex        =   24
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblOrderDate 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblCourier 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblDeliveryDate 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblDeliveryAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   8160
      TabIndex        =   20
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblDeliveryAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   19
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblDeliveryAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   8160
      TabIndex        =   18
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblDeliveryAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   17
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblDeliveryAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   16
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblDeliveryAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   15
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   14
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   11
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   10
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Advice note Order Number :-"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "frmPackaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim llngOrderNum As Long
Dim llngCustomerNum As Long
Dim mstrRoute As String
Dim mlngOrderNum As Long
Dim lstrScreenHelpFile As String
Public Property Let FindOrder(plngOrderNum As Long)

    mlngOrderNum = plngOrderNum

End Property
Public Property Get FindOrder() As Long

    FindOrder = mlngOrderNum

End Property
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property
Sub GetAcOrdNums()

    dbgOrderLinesMaster.Col = 1
    dbgOrderLinesMaster.Row = 0
    llngOrderNum = CLng(dbgOrderLinesMaster.Text)
    
    dbgOrderLinesMaster.Col = 0
    dbgOrderLinesMaster.Row = 0
    llngCustomerNum = CLng(dbgOrderLinesMaster.Text)
    
End Sub
Sub GetLocalFields()

    With gstrAdviceNoteOrder
        lblName = AmpersandDouble(Trim$(Trim$(.strSalutation) & " " & Trim$(.strInitials) & " " & Trim$(.strSurname)))
        lblAddress(0) = AmpersandDouble(Trim$(.strAdd1))
        lblAddress(1) = AmpersandDouble(Trim$(.strAdd2))
        lblAddress(2) = AmpersandDouble(Trim$(.strAdd3))
        lblAddress(3) = AmpersandDouble(Trim$(.strAdd4))
        lblAddress(4) = AmpersandDouble(Trim$(.strAdd5))
        lblAddress(5) = AmpersandDouble(Trim$(.strPostcode))
        lblDeliveryAddress(0) = AmpersandDouble(Trim$(.strDeliveryAdd1))
        lblDeliveryAddress(1) = AmpersandDouble(Trim$(.strDeliveryAdd2))
        lblDeliveryAddress(2) = AmpersandDouble(Trim$(.strDeliveryAdd3))
        lblDeliveryAddress(3) = AmpersandDouble(Trim$(.strDeliveryAdd4))
        lblDeliveryAddress(4) = AmpersandDouble(Trim$(.strDeliveryAdd5))
        lblDeliveryAddress(5) = AmpersandDouble(Trim$(.strDeliveryPostcode))
        lblCourier = Trim$(.strCourierCode)
        
        If .datCreationDate <> "00:00:00" Then
            lblDeliveryDate = Trim$(.datDeliveryDate)
            lblOrderDate = Trim$(.datCreationDate)
            If lblDeliveryDate = "00:00:00" Then
                lblDeliveryDate = "Not Specified"
            End If
        Else
            lblDeliveryDate = ""
            lblOrderDate = ""
        End If
    End With
    
End Sub
Public Sub cmdBack_Click()

    Unload Me
    ClearAdviceNote
    ClearCustomerAcount
    ClearGen
    frmAbout.Show
    
End Sub

Private Sub cmdConfirm_Click()
Dim lbooSomethingOutofStock As Boolean
Dim lstrCustName As String

    datOrderLineMaster.Refresh

    lbooSomethingOutofStock = False
    
    'if theres something out of stock
    If Not (datOrderLineMaster.Recordset.BOF = True And datOrderLineMaster.Recordset.EOF = True) Then
        Do Until datOrderLineMaster.Recordset.EOF
            If datOrderLineMaster.Recordset("Qty") > _
                datOrderLineMaster.Recordset("DespQty") Then
                
                lbooSomethingOutofStock = True
            End If
            datOrderLineMaster.Recordset.MoveNext
        Loop
    End If
    On Error GoTo 0

    CalculateRefund llngCustomerNum, llngOrderNum, lbooSomethingOutofStock
    
    cmdConfirm.Enabled = False
    dbgOrderLinesMaster.Enabled = False

    On Error Resume Next
    AddNewFileHistoryItem llngCustomerNum, llngOrderNum, lblName, "Packed"

End Sub

Private Sub cmdFind_Click()
Dim lbooCanConfirm As Boolean
Dim lstrOrderStatus As String

    If Val(txtSearchCriteria) <> 0 Then
        Busy True, Me
        
        datOrderLineMaster.RecordSource = "select * from " & gtblMasterOrderLines & " where OrderNum = " & CLng(txtSearchCriteria) & " Order by OrderLineNum, BinLocation"
        datOrderLineMaster.Refresh
        If Not (datOrderLineMaster.Recordset.BOF = True And datOrderLineMaster.Recordset.EOF = True) Then
            lbooCanConfirm = CheckCanConfirm(CLng(txtSearchCriteria), , lstrOrderStatus)
            cmdConfirm.Enabled = lbooCanConfirm
            dbgOrderLinesMaster.Enabled = lbooCanConfirm
            
            GetAcOrdNums
            GetAdviceNote llngCustomerNum, llngOrderNum
            GetLocalFields
            lblOrderStatus = lstrOrderStatus
            Busy False, Me
        Else
            Busy False, Me
            ClearAdviceNote
            GetLocalFields
            dbgOrderLinesMaster.Enabled = False
            MsgBox "No order lines found for this order!", , gconstrTitlPrefix & "Searching"
        End If
    Else
        dbgOrderLinesMaster.Enabled = False
        MsgBox "This does not appear to be an order Number!", , gconstrTitlPrefix & "Searching"
    End If
    
    lblFoundNumber = "Found " & datOrderLineMaster.Recordset.RecordCount & " records."
    
End Sub

Private Sub cmdHelp_Click()
    
    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdViewAdviceNote_Click()
Dim llngOrderNum As Long

    On Error Resume Next
    llngOrderNum = CLng(txtSearchCriteria)
    On Error GoTo 0
    If llngOrderNum = 0 Then
        MsgBox "This does not appear to be an Order Number!", , gconstrTitlPrefix & "Mandatory Field"
        Exit Sub
    End If
    
    If UCase$(App.ProductName) = "LITE" Then
        ChooseLayout ltAdviceWithAddress, Me
    Else
        ChooseLayout ltAdviceNote, Me
    End If
    
    Busy True, Me
    
    PrintObjAdviceNotesGeneral 0, 0, "S", llngOrderNum, , lblOrderStatus
    
    Busy False, Me
    ShowPlotReport
End Sub

Private Sub Form_Activate()

    If mlngOrderNum <> 0 Then
        txtSearchCriteria = mlngOrderNum
        Call cmdFind_Click
        mlngOrderNum = 0
    End If
    
End Sub
Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
        
    On Error GoTo ErrHandler
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datOrderLineMaster.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datOrderLineMaster.DatabaseName = gstrStatic.strCentralDBFile
    End Select
    
    If gstrSystemRoute <> srCompanyRoute Then
        datOrderLineMaster.Connect = gstrDBPasswords.strCentralDBPasswordString
    End If
    
    datOrderLineMaster.RecordSource = "select * from " & gtblMasterOrderLines & " where 1=0"
    datOrderLineMaster.Refresh

    ShowBanner Me, Me.Route
    
    SetupHelpFileReqs
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmPackaging.Load", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
        
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With
    
    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
    With cmdConfirm
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelp.Width + cmdHelp.Left + 120
    End With
    
    With lblFoundNumber
        .Top = (cmdHelp.Top - .Height) - 305
    End With
    
    With dbgOrderLinesMaster
        .Width = Me.Width - 360
        If (cmdHelp.Top - .Top) > 665 Then
            .Height = (cmdHelp.Top - .Top) - 665
        Else
            .Height = 665 - (cmdHelp.Top - .Top)
        End If
    End With

End Sub

Private Sub timActivity_Timer()
    
    CheckActivity
    
End Sub

Private Sub txtSearchCriteria_GotFocus()

    SetSelected Me
        
End Sub

Private Sub txtSearchCriteria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdFind_Click
    End If
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/Packing.xml>WhatsScreen"
    
    ctlBanner1.WhatsThisHelpID = IDH_PACKING_MAIN
    ctlBanner1.WhatIsID = IDH_PACKING_MAIN
    
    ctlBottomLine1.WhatsThisHelpID = IDH_PACKING_MAIN
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    cmdFind.WhatsThisHelpID = IDH_STANDARD_FIND
    cmdViewAdviceNote.WhatsThisHelpID = IDH_STANDARD_VWADNOT
    lblFoundNumber.WhatsThisHelpID = IDH_STANDARD_LBLGRITOTFOUND

    dbgOrderLinesMaster.WhatsThisHelpID = IDH_PACKING_GRIDOLM
    cmdConfirm.WhatsThisHelpID = IDH_PACKING_CONFIRM
    txtSearchCriteria.WhatsThisHelpID = IDH_PACKING_SEACRIT
    
End Sub

