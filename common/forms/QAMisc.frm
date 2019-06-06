VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmQAMisc 
   Caption         =   "Please Select an Account..."
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   ClipControls    =   0   'False
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
      TabIndex        =   17
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin MSDBGrid.DBGrid dbgAdviceNotes 
      Bindings        =   "QAMisc.frx":0000
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "QAMisc.frx":001D
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   10245
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdNotePad 
      Caption         =   "&Notepad"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdInternalNote 
      Caption         =   "&Internal Note"
      Height          =   360
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7235
      Visible         =   0   'False
      Width           =   1543
   End
   Begin VB.CommandButton cmdConsignNote 
      Caption         =   "&Consignment Note"
      Height          =   360
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7235
      Width           =   1543
   End
   Begin VB.CommandButton cmdBacthUpdate 
      Caption         =   "&Batch Status Update"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdRefund 
      Caption         =   "Place &Refund"
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7235
      Width           =   1305
   End
   Begin VB.ComboBox cboSortBy 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtSearchCriteria 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2412
   End
   Begin VB.Frame fraSearchBy 
      Caption         =   "Search By"
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   6255
      Begin VB.ComboBox cboOrderStatus 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "&Customer Number"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "CustNumber"
         Top             =   240
         Value           =   -1  'True
         Width           =   1692
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "&Order Number"
         Height          =   252
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Tag             =   "OrderNum"
         Top             =   240
         Width           =   1452
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "Customer &Name"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Tag             =   "Name"
         Top             =   480
         Width           =   1812
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "Order Status"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   7
         Tag             =   "Status"
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Data datAdviceNotes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   10440
      Top             =   3960
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   21
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin VB.Label lblFoundNumber 
      Caption         =   "Found 0 records"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6675
      Width           =   2415
   End
   Begin VB.Label lblSortby 
      Alignment       =   1  'Right Justify
      Caption         =   "Sort By"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Criteria :-"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmQAMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrOrderStatus() As String
Dim lstrExtraSQL As String
Dim mstrRoute As String
Dim lstrScreenHelpFile As String
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property
Private Sub cboOrderStatus_Click()

    optSearchField(3).Value = True
    optSearchField_Click (3)
    
End Sub
Public Sub cmdBack_Click()

    Unload Me
    frmAbout.Show
    
End Sub
Sub GetAdviceNoteData()
    
End Sub
Private Sub cmdBacthUpdate_Click()
Dim lintRetVal As Integer
Dim lintRetVal2 As Integer
Dim lstrDespacthDate As Date

    If datAdviceNotes.Recordset.RecordCount <= 0 Then
        MsgBox "You must have some records displayed in the grid!", , gconstrTitlPrefix & "Batch Update"
        Exit Sub
    End If
    
    frmChildOptions.List = "Order Status"
    
    If frmChildOptions.List <> "" Then
        frmChildOptions.Code = "C"
        frmChildOptions.Show vbModal
    Else
        Exit Sub
    End If
    
    If Trim(frmChildOptions.Code) <> "" Then
        lintRetVal = MsgBox("You are about to update all of the order shown" & vbCrLf & _
            "in the grid with Order Status '" & Trim$(frmChildOptions.Code) & "'." & vbCrLf & vbCrLf & _
            "Click YES to proceed!", vbYesNo, gconstrTitlPrefix & "Batch Order Status Update")
        If lintRetVal = vbYes Then
            lintRetVal2 = MsgBox("Would you also like to update the Despatch date?", vbYesNo, gconstrTitlPrefix & "Batch Order Status Update")
                If lintRetVal2 = vbYes Then
                
                    frmChildCalendar.CalDate = Now()
                    frmChildCalendar.Show vbModal
                    lstrDespacthDate = frmChildCalendar.CalDate
                    BatchUpdateOrderStatus lstrExtraSQL, Trim(frmChildOptions.Code), lstrDespacthDate, "DESPDATE"
                Else
                    BatchUpdateOrderStatus lstrExtraSQL, Trim(frmChildOptions.Code), 0
                End If
        End If
    End If
    
    datAdviceNotes.Refresh
    
End Sub
Private Sub cmdFind_Click()
Const lstrEndOfMessage = ", or select a different search method"

    lstrExtraSQL = ""
    If optSearchField(0).Value = True Then ' Customer Number
        If CLng(Val(txtSearchCriteria)) = 0 Then
            MsgBox "You must enter a Customer Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where CustNum = " & CLng(txtSearchCriteria)
        End If
    ElseIf optSearchField(2).Value = True Then ' Order Number
        If CLng(Val(txtSearchCriteria)) = 0 Then
            MsgBox "You must enter a Order Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where OrderNum = " & CLng(txtSearchCriteria)
        End If
    ElseIf optSearchField(1).Value = True Then ' Customer name
        If Trim$(txtSearchCriteria) = "" Then
            MsgBox "You must enter a Customer Name" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where CallerSurname = '" & Trim$(txtSearchCriteria) & "'"
        End If
    ElseIf optSearchField(3).Value = True Then ' Order Status
        If Trim$(cboOrderStatus) = "" Then
            MsgBox "You must select a Status" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
        
            lstrExtraSQL = "where orderStatus = '" & Trim$(NotNull(cboOrderStatus, lstrOrderStatus)) & "'"
        End If
    End If

    Busy True, Me
    
    datAdviceNotes.RecordSource = "SELECT OrderStatus, AuthorisationCode, OrderNum, CardNumber, " & _
        "ExpiryDate, OrderType, PaymentType2, CallerSalutation, CallerSurname, CallerInitials, " & _
        "AdviceAdd1, AdviceAdd2, AdviceAdd3, AdviceAdd4, AdviceAdd5, AdvicePostcode, CardType, " & _
        "CardIssueNumber, CardStartDate, Trim(Format([Denom],'0.00')) & Format([Donation],'0.00') " & _
        "AS Donat, Trim(Format([Denom],'0.00')) & Format([Payment],'0.00') AS Pay1, " & _
        "Trim(Format([Denom],'0.00')) & Format([Payment2],'0.00') AS Pay2, " & _
        "Trim(Format([Denom],'0.00')) & Format([Underpayment],'0.00') AS UndP, " & _
        "Trim(Format([Denom],'0.00')) & Format([Reconcilliation],'0.00') AS Recon, " & _
        "Trim(Format([Denom],'0.00')) & Format([Postage],'0.00') AS Post, Trim(Format([Denom],'0.00')) " & _
        "& Format([Vat],'0.00') AS TaxVat, Trim(Format([Denom],'0.00')) & Format([TotalIncVat],'0.00') " & _
        "AS Total, ProcessedBy, CreationDate, CustNum from " & gtblAdviceNotes & " " & lstrExtraSQL & " order by " & cboSortBy & ";"
    
    datAdviceNotes.Refresh

    lblFoundNumber = "Found " & datAdviceNotes.Recordset.RecordCount & " records."
    Busy False, Me

End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdInternalNote_Click()
Dim lstrLockingFlag As String
    
    lstrLockingFlag = LockingPhaseGen(True)
    
    gstrInternalNote.strType = "Internal"

    On Error Resume Next
    gstrAdviceNoteOrder.lngOrderNum = CLng(datAdviceNotes.Recordset("OrderNum"))
    gstrAdviceNoteOrder.lngAdviceRemarkNum = CLng(datAdviceNotes.Recordset("AdviceRemarkNum"))

    If Err.Number = 3021 Then
        MsgBox "You must select an order!", , gconstrTitlPrefix & "Internal Note"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    If gstrAdviceNoteOrder.lngAdviceRemarkNum <> 0 Then
        If gstrInternalNote.lngRemarkNumber <> 0 Then
            frmChildNote.NoteText = gstrInternalNote.strText
            frmChildNote.NoteType = "Internal Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        Else
            gstrInternalNote.lngRemarkNumber = gstrAdviceNoteOrder.lngAdviceRemarkNum
            GetRemark gstrInternalNote.lngRemarkNumber, gstrInternalNote
            frmChildNote.NoteText = gstrInternalNote.strText
            frmChildNote.NoteType = "Internal Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        End If
    Else
        AddNewRemark lstrLockingFlag
        GetRemarkNum lstrLockingFlag, gstrInternalNote
        gstrAdviceNoteOrder.lngAdviceRemarkNum = gstrInternalNote.lngRemarkNumber
        UpdateRemarkAdviceID gstrAdviceNoteOrder.lngOrderNum, gstrAdviceNoteOrder.lngAdviceRemarkNum, "Internal"
        frmChildNote.NoteText = gstrInternalNote.strText
        frmChildNote.NoteType = "Internal Note Comments"
        Load frmChildNote
        frmChildNote.Show vbModal
    End If
    
    gstrInternalNote.strText = frmChildNote.NoteText
    gstrInternalNote.strType = frmChildNote.NoteType
    UpdateRemark gstrAdviceNoteOrder.lngAdviceRemarkNum, gstrInternalNote.strType, gstrInternalNote.strText

    gstrAdviceNoteOrder.lngAdviceRemarkNum = 0
    gstrInternalNote.lngRemarkNumber = 0
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmQAMisc.cmdInternalNote_Click", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Sub
Private Sub cmdConsignNote_Click()
Dim lstrLockingFlag As String

    lstrLockingFlag = LockingPhaseGen(True)
    
    gstrConsignmentNote.strType = "Consignment"

    On Error Resume Next
    gstrAdviceNoteOrder.lngOrderNum = CLng(datAdviceNotes.Recordset("OrderNum"))
    gstrAdviceNoteOrder.lngConsignRemarkNum = CLng(datAdviceNotes.Recordset("ConsignRemarkNum"))

    If Err.Number = 3021 Then
        MsgBox "You must select an order!", , gconstrTitlPrefix & "Consignment Note"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 Then
        If gstrConsignmentNote.lngRemarkNumber <> 0 Then
            frmChildNote.NoteText = gstrConsignmentNote.strText
            frmChildNote.NoteType = "Consignment Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        Else
            gstrConsignmentNote.lngRemarkNumber = gstrAdviceNoteOrder.lngConsignRemarkNum
            GetRemark gstrConsignmentNote.lngRemarkNumber, gstrConsignmentNote
            gstrAdviceNoteOrder.lngConsignRemarkNum = gstrConsignmentNote.lngRemarkNumber
            frmChildNote.NoteText = gstrConsignmentNote.strText
            frmChildNote.NoteType = "Consignment Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        End If
    Else
        AddNewRemark lstrLockingFlag
        GetRemarkNum lstrLockingFlag, gstrConsignmentNote
        gstrAdviceNoteOrder.lngConsignRemarkNum = gstrConsignmentNote.lngRemarkNumber
        UpdateRemarkAdviceID gstrAdviceNoteOrder.lngOrderNum, gstrAdviceNoteOrder.lngConsignRemarkNum, "Consignment"
        ToggleRemarkInUseBy gstrConsignmentNote.lngRemarkNumber, False
        frmChildNote.NoteText = gstrConsignmentNote.strText
        frmChildNote.NoteType = "Consignment Note Comments"
        Load frmChildNote
        frmChildNote.Show vbModal
    End If
    
    gstrConsignmentNote.strText = frmChildNote.NoteText
    gstrConsignmentNote.strType = frmChildNote.NoteType
    UpdateRemark gstrAdviceNoteOrder.lngConsignRemarkNum, gstrConsignmentNote.strType, gstrConsignmentNote.strText

    gstrAdviceNoteOrder.lngConsignRemarkNum = 0
    gstrConsignmentNote.lngRemarkNumber = 0

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmQAMisc.cmdConsignNote_Click", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Sub
Private Sub cmdNotePad_Click()

    frmChildCuNotes.Show vbModal
    
End Sub

Private Sub cmdRefund_Click()
Dim lstrInUseByFlag As String
Dim llngOrderNum As Long
Dim lstrRefundValue As String
Dim lstrSQL As String
Dim lstrName As String
Dim lintRetVal As Variant
Dim lstrLockingFlag As String
Dim lstrOrderStatus As String

    On Error Resume Next
    lstrOrderStatus = (datAdviceNotes.Recordset("OrderStatus"))
    If Err.Number = 3021 Then
        MsgBox "You must select an order!", , gconstrTitlPrefix & "Refund"
        Exit Sub
    End If
            
    If Trim$(lstrOrderStatus) = "X" Or Trim$(lstrOrderStatus) = "R" Then
        MsgBox "You may not place a Refund against this order!", , gconstrTitlPrefix & "Refund"
        Exit Sub
    End If
    
    NewRefund
    Exit Sub

    lstrLockingFlag = LockingPhaseGen(True)

    gstrConsignmentNote.strType = "Consignment"
    
    With gstrAdviceNoteOrder
        
        On Error Resume Next
        .lngCustNum = CLng(datAdviceNotes.Recordset("CustNum"))
        llngOrderNum = CLng(datAdviceNotes.Recordset("OrderNum"))
           
        If Err.Number = 3021 Then
            MsgBox "You must select an order!", , gconstrTitlPrefix & "Refund"
            Exit Sub
        End If
        On Error GoTo ErrHandler
        
        If .lngCustNum = 0 Or llngOrderNum = 0 Then
            MsgBox "You must select an order!", , gconstrTitlPrefix & "Refund"
            Exit Sub
        End If
        
        lintRetVal = MsgBox("Do you wish to create a Refund for " & vbCrLf & _
            "Customer Number M" & .lngCustNum & ", Order Number " & llngOrderNum & " ?", vbYesNo + vbDefaultButton1, gconstrTitlPrefix & "Refund")
        
        If lintRetVal = vbNo Then
            Exit Sub
        End If
        
        lstrRefundValue = SystemPrice(PriceVal(InputBox("Please enter the Amount to be refunded, eg. �25.50", "Refund Amount")))
        
        If PriceVal(lstrRefundValue) = 0 Then
            MsgBox "You must enter a valid value", , gconstrTitlPrefix & "Refund"
            Exit Sub
        End If
        
        lstrInUseByFlag = LockingPhaseGen(True)

        .lngCustNum = CLng(datAdviceNotes.Recordset("CustNum"))
        llngOrderNum = CLng(datAdviceNotes.Recordset("OrderNum"))
        GetAdviceNote .lngCustNum, llngOrderNum
    
       'update some of the values
       
        .datBankRepPrintDate = CDate("0")
        .datCreationDate = CDate(Now())
        .datDeliveryDate = CDate("0")
        .datDespatchDate = CDate("0")
        .intNumOfParcels = 0
        .lngAdviceRemarkNum = 0
        .lngConsignRemarkNum = 0
        .lngGrossWeight = 0
        .strAcctInUseByFlag = ""
        .strAuthorisationCode = ""
        .strDonation = "0"
        .strOOSRefund = "0"
        .strOrderCode = "O"
        .strOrderStyle = "3"
        .strPayment = "0"
        .strPayment2 = "0"
        .strPaymentType1 = "V"
        .strPaymentType2 = ""
        .strPostage = "0"
        .strReconcilliation = CCur(lstrRefundValue)
        .strTotalIncVat = "0"
        .strUnderpayment = "0"
        .strVAT = "0"
        .strCardName = ""
        .strCardNumber = ""
        .datExpiryDate = CDate("0")
        .datCardStartDate = CDate("0")
        
        gstrCustomerAccount.lngCustNum = CLng(datAdviceNotes.Recordset("CustNum"))
        GetCustomerAccount .lngCustNum, False
        lstrName = Trim$(Trim$(.strSalutation) & " " & Trim$(.strInitials) & " " & Trim$(.strSurname))
    
         AddAdviceNote lstrInUseByFlag, "C"
         GetAdviceOrderNum lstrInUseByFlag, .lngCustNum
         UpdateAdviceNote
         UpdateOrderStatus "C", 0, "S", .lngOrderNum 'J.M. Requested 10/08/00
         MsgBox "Your Refunded Order number is " & .lngOrderNum & " and will be available in your normal Refund Advice note print!", , gconstrTitlPrefix & "Refund"
         UpdateSalesCode
        
         lstrSQL = "INSERT INTO " & gtblMasterOrderLines & " ( CustNum, OrderNum, " & _
            "CatNum, ItemDescription, BinLocation, Qty, DespQty, Price, " & _
            "Vat, Weight, TaxCode, TotalPrice, TotalWeight, Class, SalesCode, Denom ) " & _
            "SELECT " & .lngCustNum & " AS Expr1, " & .lngOrderNum & _
            " AS Expr2, 'REFUND' AS Expr3, " & _
            "'Refund' AS Expr4, ' ' AS Expr5, 0 AS Expr6, " & _
            "0 AS Expr7, 0 AS Expr8, 0 AS Expr9, 0 AS Expr10, " & _
            "'Z' AS Expr11, 0 AS Expr12, 0 AS Expr13, " & _
            "0 AS Expr14, 0 AS Expr15, '" & gstrReferenceInfo.strDenomination & "' as De;"

        gdatCentralDatabase.Execute lstrSQL

        AddNewRemark lstrLockingFlag
        GetRemarkNum lstrLockingFlag, gstrConsignmentNote
        gstrAdviceNoteOrder.lngConsignRemarkNum = gstrConsignmentNote.lngRemarkNumber
        UpdateRemarkAdviceID gstrAdviceNoteOrder.lngOrderNum, gstrAdviceNoteOrder.lngConsignRemarkNum, "Consignment"
        ToggleRemarkInUseBy gstrConsignmentNote.lngRemarkNumber, False
        frmChildNote.NoteText = Trim(gstrConsignmentNote.strText)
        frmChildNote.NoteType = "Consignment Note Comments"
        Load frmChildNote
        frmChildNote.Show vbModal
        
        gstrConsignmentNote.strText = frmChildNote.NoteText
        gstrConsignmentNote.strType = frmChildNote.NoteType
        UpdateRemark gstrAdviceNoteOrder.lngConsignRemarkNum, gstrConsignmentNote.strType, gstrConsignmentNote.strText
    
        gstrAdviceNoteOrder.lngConsignRemarkNum = 0
        gstrConsignmentNote.lngRemarkNumber = 0
                 
        'add cheque
        AddCashBookEntry "", CCur(lstrRefundValue), lstrName, .lngCustNum, .lngOrderNum, "REFUND"
         
        ToggleAdviceInUseBy gstrAdviceNoteOrder.lngOrderNum, False
         
        AddNewFileHistoryItem .lngCustNum, .lngOrderNum, lstrName, "Refunded"
         
        ClearCustomerAcount
        ClearAdviceNote
        ClearGen
    End With

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmQAMisc.cmdRefund_Click", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Sub

Private Sub dbgAdviceNotes_ButtonClick(ByVal ColIndex As Integer)

    frmChildOptions.List = ""
    
    Select Case ColIndex
    Case 0  'Order Status
        frmChildOptions.List = "Order Status"
    Case 5, 6   'PaymentType 1 & 2
        frmChildOptions.List = "Payment Method"
    Case 19 'Card Type
        frmChildOptions.List = "Credit Card Type"
    End Select
    
    If frmChildOptions.List <> "" Then
        frmChildOptions.Code = dbgAdviceNotes.Columns(ColIndex).Value
        frmChildOptions.Show vbModal
        dbgAdviceNotes.Columns(ColIndex).Value = frmChildOptions.Code
    End If

End Sub

Private Sub Form_Activate()

    dbgAdviceNotes.Visible = True
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    If gbooSQLServerInUse = True Then
        datAdviceNotes.Connect = "ODBC;DATABASE=Mmos;DSN=Mmos;"
    
    Else
        Select Case gstrUserMode
        Case gconstrTestingMode
            datAdviceNotes.DatabaseName = gstrStatic.strCentralTestingDBFile
        Case gconstrLiveMode
            datAdviceNotes.DatabaseName = gstrStatic.strCentralDBFile
        End Select
        
        If gstrSystemRoute <> srCompanyRoute Then
            datAdviceNotes.Connect = gstrDBPasswords.strCentralDBPasswordString
        End If
    End If
    
    FillList "Order Status", cboOrderStatus, lstrOrderStatus()
    FillFieldSortBy
    cboSortBy.ListIndex = 0
    
    datAdviceNotes.RecordSource = "select * from " & gtblAdviceNotes & " " & _
        "where 1=0" ' order by orderNum"

    If gstrReferenceInfo.booDonationAvail = False Then
        dbgAdviceNotes.Columns(22).Visible = False
    End If
    
    ShowBanner Me, Me.Route
    
    SetupHelpFileReqs
    
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
    
    With cmdNotePad
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelp.Width + cmdHelp.Left + 120
    End With
    
    With cmdRefund
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdNotePad.Width + cmdNotePad.Left + 120
    End With
    
    With cmdConsignNote
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdRefund.Width + cmdRefund.Left + 120
    End With
    
    With cmdInternalNote
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdConsignNote.Width + cmdConsignNote.Left + 120
    End With
    
    With lblFoundNumber
        .Top = (cmdHelp.Top - .Height) - 305
    End With
    
    With dbgAdviceNotes
        .Width = Me.Width - 360 '360 '240
        If (cmdHelp.Top - .Top) > 665 Then
            .Height = (cmdHelp.Top - .Top) - 665
        Else
            .Height = 665 - (cmdHelp.Top - .Top)
        End If
    End With

End Sub

Private Sub optSearchField_Click(Index As Integer)

    Select Case Index
    Case 3
        cboOrderStatus.Enabled = True
        cboOrderStatus.BackColor = vbWindowBackground
        txtSearchCriteria.Enabled = False
        txtSearchCriteria.BackColor = vbActiveBorder
    Case Else
        cboOrderStatus.Enabled = False
        cboOrderStatus.BackColor = vbActiveBorder
        txtSearchCriteria.Enabled = True
        txtSearchCriteria.BackColor = vbWindowBackground
    End Select
    
End Sub
Sub FillFieldSortBy()
Dim tblAdviceNotes As TableDef
Dim lintArrInc As Integer

    Set tblAdviceNotes = gdatCentralDatabase.TableDefs(gtblAdviceNotes)

    For lintArrInc = 0 To tblAdviceNotes.Fields.Count - 1
        cboSortBy.AddItem tblAdviceNotes.Fields(lintArrInc).Name
    Next lintArrInc
    
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
Sub BatchUpdateOrderStatus(pstrWhereClause As String, pstrStatus As String, _
    pdatDespDate As Date, Optional pstrParam As Variant)
Dim lstrSQL As String
    
    On Error GoTo ErrHandler

    If IsMissing(pstrParam) = True Then
        pstrParam = ""
    End If
    
    Select Case pstrParam
    Case "DESPDATE"
        lstrSQL = "UPDATE " & gtblAdviceNotes & " SET OrderStatus = '" & pstrStatus & "' "
        lstrSQL = lstrSQL & ", " & gtblAdviceNotes & ".DespatchDate = '" & pdatDespDate & "' " & pstrWhereClause
    Case Else
        lstrSQL = "UPDATE " & gtblAdviceNotes & " SET OrderStatus = '" & pstrStatus & "' " & pstrWhereClause
    End Select
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "BatchUpdateOrderStatus", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/OrderMaintenance.xml>WhatsScreen"
    
    ctlBanner1.WhatsThisHelpID = IDH_ORDMNT_MAIN
    ctlBanner1.WhatIsID = IDH_ORDMNT_MAIN
    
    ctlBottomLine1.WhatsThisHelpID = IDH_ORDMNT_MAIN
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    cmdFind.WhatsThisHelpID = IDH_STANDARD_FIND
    cmdNotePad.WhatsThisHelpID = IDH_STANDARD_CUNOTES
    cmdConsignNote.WhatsThisHelpID = IDH_STANDARD_CONSNOTE
    
    lblFoundNumber.WhatsThisHelpID = IDH_STANDARD_LBLGRITOTFOUND
    
    fraSearchBy.WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(0).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(1).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(2).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(3).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    cboOrderStatus.WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    
    cmdRefund.WhatsThisHelpID = IDH_ORDMNT_REFUND
    lblSortby.WhatsThisHelpID = IDH_ORDMNT_SORTBY
    cboSortBy.WhatsThisHelpID = IDH_ORDMNT_SORTBY
    dbgAdviceNotes.WhatsThisHelpID = IDH_ORDMNT_GRIDADV
    txtSearchCriteria.WhatsThisHelpID = IDH_ORDMNT_SEARCHCRIT
    
End Sub
Sub NewRefund()
Dim llngOrderNum As Long
Dim llngCustNum As Long
Dim lintRetVal As Integer

    On Error Resume Next
    llngCustNum = CLng(datAdviceNotes.Recordset("CustNum"))
    llngOrderNum = CLng(datAdviceNotes.Recordset("OrderNum"))
    
    If Err.Number = 3021 Then
        MsgBox "You must select an order!", , gconstrTitlPrefix & "Refund"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    If llngCustNum = 0 Or llngOrderNum = 0 Then
        MsgBox "You must select an order!", , gconstrTitlPrefix & "Refund"
        Exit Sub
    End If
    
    frmChildRefundSel.CustNum = llngCustNum
    frmChildRefundSel.OrderNum = llngOrderNum
    frmChildRefundSel.Show vbModal
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "NewRefund", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
End Sub
