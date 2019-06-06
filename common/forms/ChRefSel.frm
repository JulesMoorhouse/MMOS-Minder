VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmChildRefundSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Refund Selections"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh Grid"
      Height          =   360
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Caption         =   "Refund Details"
      Height          =   1695
      Left            =   5640
      TabIndex        =   30
      Top             =   4800
      Width           =   4695
      Begin VB.TextBox txtRefundPayee 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtRefundPostage 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkRefundPostage 
         Alignment       =   1  'Right Justify
         Caption         =   "Refund Postage?"
         Height          =   195
         Left            =   45
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboPaymentType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   3012
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Refund Payee"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Postage"
         Height          =   255
         Left            =   2400
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Refund"
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblTotalRefund 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblPaymentType 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Type"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Original Order Totals"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   5415
      Begin VB.Label lblTotalIncVat 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblPostage 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   28
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblVAT 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblUnderPayment 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblReconcilliation 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblDonation 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblPayment2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblPayment1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblDonationCaption 
         Alignment       =   1  'Right Justify
         Caption         =   "Donation"
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Under Payment"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Reconciliation"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Inc VAT"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Postage"
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "VAT"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblPayment1Caption 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Payment 1"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblPayment2Caption 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Payment 2"
         Height          =   255
         Left            =   15
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   360
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1305
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   360
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save"
      Height          =   360
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   1305
   End
   Begin VB.Data datRefundOrderLines 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data datOriginalOrderLines 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgRefundOrderLines 
      Bindings        =   "ChRefSel.frx":0000
      CausesValidation=   0   'False
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "ChRefSel.frx":0022
      TabIndex        =   4
      Top             =   2640
      Width           =   10215
   End
   Begin MSDBGrid.DBGrid dbgOriginalOrderLines 
      Bindings        =   "ChRefSel.frx":1425
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "ChRefSel.frx":1449
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   35
      Top             =   7275
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13229
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "14/08/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "17:24"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   6600
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label lblReconWarning 
      Caption         =   "Any reconciliation amount will have already been dealt with."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7440
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "You may modify the prices in the grid below as required"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frmChildRefundSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrScreenHelpFile As String
Dim mlngCustNum As Long
Dim mlngOrderNum As Long
Dim lstrPaymentTypeCode() As String
Public Property Get OrderNum() As Long

    OrderNum = mlngOrderNum
    
End Property
Public Property Let OrderNum(plngOrderNum As Long)

    mlngOrderNum = plngOrderNum

End Property
Public Property Get CustNum() As Long

    CustNum = mlngCustNum
    
End Property
Public Property Let CustNum(plngCustNum As Long)

    mlngCustNum = plngCustNum

End Property

Private Sub chkRefundPostage_Click()

    If chkRefundPostage.Value = vbChecked Then
        txtRefundPostage = SystemPrice(lblPostage)
    Else
        txtRefundPostage = SystemPrice(0)
    End If
    
    UpdateTotalRefund
    
End Sub

Private Sub cmdAdd_Click()
Dim lstrCatNum As String
Dim lstrSQL As String

    On Error Resume Next
    lstrCatNum = Trim$(datOriginalOrderLines.Recordset("CatNum"))
    If lstrCatNum = "" Then
        MsgBox "You must select a product from the top grid, if you wish to add it to the refund order!", vbInformation, gconstrTitlPrefix & "Refund Process"
        Exit Sub
    End If
    
    lstrSQL = "INSERT INTO " & gtblOrderLines & " ( CustNum, CatNum, ItemDescription, BinLocation, " & _
        "Qty, Price, Vat, Weight, TaxCode, TotalPrice, TotalWeight, Class, SalesCode, " & _
        "OrderLineNum, ParcelNumber ) SELECT " & gtblMasterOrderLines & ".CustNum, " & _
        gtblMasterOrderLines & ".CatNum, " & gtblMasterOrderLines & _
        ".ItemDescription, " & gtblMasterOrderLines & ".BinLocation, " & gtblMasterOrderLines & _
        ".Qty, " & gtblMasterOrderLines & ".Price, " & gtblMasterOrderLines & _
        ".Vat, " & gtblMasterOrderLines & ".Weight, " & gtblMasterOrderLines & _
        ".TaxCode, " & gtblMasterOrderLines & ".TotalPrice, " & _
        gtblMasterOrderLines & ".TotalWeight, " & gtblMasterOrderLines & _
        ".Class, " & gtblMasterOrderLines & ".SalesCode, " & gtblMasterOrderLines & _
        ".OrderLineNum, " & gtblMasterOrderLines & ".ParcelNumber FROM " & _
        gtblMasterOrderLines & " WHERE trim(" & gtblMasterOrderLines & ".CatNum='" & _
        lstrCatNum & "') AND " & gtblMasterOrderLines & ".OrderNum=" & mlngOrderNum & ";"
        
    gdatLocalDatabase.Execute lstrSQL
    datRefundOrderLines.Refresh
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()
Dim lstrLockingFlag As String
Dim lstrInUseByFlag As String
Dim lstrName As String
Dim lstrSQL As String

    If cboPaymentType = "" Then
        MsgBox "You must specific how this refund is to be paid!", vbInformation, gconstrTitlPrefix & "Refund Process"
        Exit Sub
    End If
    
    'Copy original order
    With gstrAdviceNoteOrder
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
        .strOrderStyle = "1" ' regular
        .strUnderpayment = "0"
        .strVAT = SystemPrice(CCur("-" & Val(gstrVatTotal))) '"0"
        .strPostage = CCur("-" & CCur(txtRefundPostage))
        If Trim$(UCase$(cboPaymentType)) = "CREDIT CARD" Then
            'Copy Credit card number etc
            'already in the buffer!
        Else
            .strCardName = txtRefundPayee
            .strCardNumber = ""
            .datExpiryDate = CDate("0")
            .datCardStartDate = CDate("0")
            .lngIssueNumber = 0
        End If
    
        .strPaymentType1 = Trim$(NotNull(cboPaymentType, lstrPaymentTypeCode))
            
        .strOrderCode = "O"  ' Post
        .strPayment = CCur("-" & Val(CCur(lblTotalRefund)))
        .strTotalIncVat = CCur("-" & Val(CCur(lblTotalRefund)))
         
        lstrInUseByFlag = LockingPhaseGen(True)
        gstrCustomerAccount.lngCustNum = mlngCustNum
        GetCustomerAccount mlngCustNum, False
        
        AddAdviceNote lstrInUseByFlag, "R"
        GetAdviceOrderNum lstrInUseByFlag, .lngCustNum
        UpdateOrderStatus "R", 0, "S", .lngOrderNum
        
        'Add to new refund fields
        lstrSQL = "UPDATE " & gtblAdviceNotes & " SET RefundOrignNum = " & mlngOrderNum & ", " & _
            "RefundReason = 'REFUND', ChequeRequestDate = #" & Format(Date, "dd/mmm/yyyy") & "# " & _
            "WHERE CustNum=" & mlngCustNum & " AND OrderNum=" & .lngOrderNum & ";"
        gdatCentralDatabase.Execute lstrSQL
         
         MsgBox "Your Refunded Order number is " & .lngOrderNum & " and will be available in your normal Refund Advice note print!", , gconstrTitlPrefix & "Refund"
         UpdateSalesCode
        
        'add refund product
        AppendOrderLinesToMaster mlngCustNum, gstrAdviceNoteOrder.lngOrderNum
        
        'Add Consignment note
        lstrLockingFlag = LockingPhaseGen(True)
        AddNewRemark lstrLockingFlag
        GetRemarkNum lstrLockingFlag, gstrConsignmentNote
        gstrAdviceNoteOrder.lngConsignRemarkNum = gstrConsignmentNote.lngRemarkNumber
        UpdateRemarkAdviceID gstrAdviceNoteOrder.lngOrderNum, gstrAdviceNoteOrder.lngConsignRemarkNum, "Consignment"
        ToggleRemarkInUseBy gstrConsignmentNote.lngRemarkNumber, False

        frmChildNote.NoteText = "This refund relates to previous order, numbered " & mlngOrderNum & "."
        frmChildNote.NoteType = "Consignment Note Comments"
        Load frmChildNote
        frmChildNote.Show vbModal
        
        gstrConsignmentNote.strText = frmChildNote.NoteText
        gstrConsignmentNote.strType = "Consignment"
        UpdateRemark gstrAdviceNoteOrder.lngConsignRemarkNum, gstrConsignmentNote.strType, gstrConsignmentNote.strText
    
        UpdateAdviceNote
         
        ToggleAdviceInUseBy gstrAdviceNoteOrder.lngOrderNum, False
        
        lstrName = Trim$(Trim$(.strSalutation) & " " & Trim$(.strInitials) & " " & Trim$(.strSurname))
        
        AddNewFileHistoryItem .lngCustNum, .lngOrderNum, lstrName, "Refunded"
         
        ClearCustomerAcount
        ClearAdviceNote
        ClearGen
    End With
    
    Unload Me
    
End Sub
Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)

End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub
Private Sub cmdRefresh_Click()

    UpdateOrderLinesTotals
    UpdateTotalRefund
    datRefundOrderLines.Refresh
    dbgRefundOrderLines.Refresh
    
End Sub

Private Sub cmdRemove_Click()
Dim lstrCatNum As String
Dim lstrSQL As String

    On Error Resume Next
    lstrCatNum = Trim$(datRefundOrderLines.Recordset("CatNum"))
    If lstrCatNum = "" Then
        MsgBox "You must select a product from the bottom grid, if you wish to remove it from the refund order!", vbInformation, gconstrTitlPrefix & "Refund Process"
        Exit Sub
    End If
    
    lstrSQL = "Delete * from " & gtblOrderLines & " where CustNum=" & mlngCustNum & " and CatNum='" & lstrCatNum & "';"
    gdatLocalDatabase.Execute lstrSQL
    
    datRefundOrderLines.Refresh
        
End Sub

Private Sub dbgRefundOrderLines_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    UpdateOrderLinesTotals
    UpdateTotalRefund
    dbgRefundOrderLines.Refresh
    
End Sub

Private Sub Form_Load()
Dim lstrSQL As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    If mlngCustNum = 0 Or mlngOrderNum = 0 Then
        Exit Sub
    Else
        GetAdviceNote mlngCustNum, mlngOrderNum
    End If
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datRefundOrderLines.DatabaseName = gstrStatic.strLocalTestingDBFile
        datOriginalOrderLines.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datOriginalOrderLines.DatabaseName = gstrStatic.strCentralDBFile
        datRefundOrderLines.DatabaseName = gstrStatic.strLocalDBFile
    End Select
    
    If gstrSystemRoute <> srCompanyRoute Then
        datOriginalOrderLines.Connect = gstrDBPasswords.strCentralDBPasswordString
        datRefundOrderLines.Connect = gstrDBPasswords.strLocalDBPasswordString
    End If

    lstrSQL = "Delete * from " & gtblOrderLines & ";"
    gdatLocalDatabase.Execute lstrSQL
    
    lstrSQL = "Select CatNum, ItemDescription as Description, Qty, DespQty, " & _
        "trim(Denom) & format(Price,'0.00') as UnitPrice, " & _
        "trim(Denom) & format(Vat,'0.00') as Vat2, TaxCode, " & _
        "trim(Denom) & format(TotalPrice,'0.00') as TP, " & _
        "OrderLineNum as LineNum, ParcelNumber " & _
        " from " & gtblMasterOrderLines & " where CustNum=" & mlngCustNum & _
        " and OrderNum=" & mlngOrderNum & _
        " order by OrderLineNum, OrderNum"
                        
                        
    datOriginalOrderLines.RecordSource = lstrSQL
    datRefundOrderLines.RecordSource = "Select * from " & gtblOrderLines & _
        " where CustNum=" & mlngCustNum & " order by OrderLineNum"
    
    FillList "Payment Method", cboPaymentType, lstrPaymentTypeCode()
    
    GetLocalFields
    
    SetupHelpFileReqs
    
End Sub
Sub GetLocalFields()

    With gstrAdviceNoteOrder
        lblDonation = SystemPrice(.strDonation) & " "
        lblPayment1 = SystemPrice(.strPayment) & " "
        lblPayment2 = SystemPrice(.strPayment2) & " "
        lblPostage = SystemPrice(.strPostage) & " "
        lblVAT = SystemPrice(.strVAT) & " "
        
        lblTotalIncVat = SystemPrice(.strTotalIncVat) & " "
        lblUnderPayment = SystemPrice(.strUnderpayment) & " "
        lblReconcilliation = SystemPrice(.strReconcilliation) & " "
        
        If Val(CCur(lblReconcilliation)) > 0 Then
            lblReconWarning.Visible = True
        End If
        
        lblPayment1Caption = "Payment (" & GetListCodeDesc("Payment Method", .strPaymentType1) & ")"
        If lblPayment1Caption = "Payment ()" Then
            lblPayment1Caption.Visible = False
            lblPayment1.Visible = False
        Else
            lblPayment1Caption.Visible = True
            lblPayment1.Visible = True
        End If
        
        lblPayment2Caption = "Payment (" & GetListCodeDesc("Payment Method", .strPaymentType2) & ")"
        If lblPayment2Caption = "Payment ()" Then
            lblPayment2Caption.Visible = False
            lblPayment2.Visible = False
        Else
            lblPayment2Caption.Visible = True
            lblPayment2.Visible = True
        End If
        
         If Trim$(.strCardName) <> "" Or Not IsBlank(.strCardName) Then
            txtRefundPayee = Trim$(Trim$(.strSalutation) & " " & Trim$(.strInitials) & " " & Trim$(.strSurname))
         End If
    End With
    
End Sub
Sub UpdateTotalRefund()
Dim lintOrderSubTotal As Integer
Dim lbooPostageSetHere As Boolean
Dim lstrPostageCode As String
Dim lstrPostageListVars As ListVars
Dim lcurVAT As Currency
Dim lcurOrderTotal As Currency

    OrderTotal mlngCustNum
    lcurVAT = SystemPrice(gstrVatTotal)
    lcurOrderTotal = SystemPrice(gstrOrderTotal)
    
    If chkRefundPostage.Value = vbChecked Then
        lblTotalRefund = SystemPrice(CCur(lcurOrderTotal) + CCur(lblPostage))
    Else
        lblTotalRefund = SystemPrice(CCur(lcurOrderTotal))
    End If
    
    If CCur(lblPayment1) + CCur(lblPayment2) < CCur(lblTotalRefund) Then
        lblTotalRefund = SystemPrice(CCur(lblPayment1) + CCur(lblPayment2))
        lblWarning = "The customer paid less than the sum of the refund!"
        lblWarning.Visible = True
    Else
        lblWarning.Visible = False
    End If
    
End Sub

Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
        
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP

End Sub

Private Sub txtRefundPayee_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtRefundPayee_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub

Private Sub txtRefundPostage_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtRefundPostage_LostFocus()

    txtRefundPostage = SystemPrice(txtRefundPostage)
    UpdateTotalRefund
    
End Sub
