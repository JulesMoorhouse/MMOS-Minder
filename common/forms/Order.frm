VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmOrder 
   Caption         =   "finalise order"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   285
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
      TabIndex        =   19
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.CommandButton cmdOrderStatus 
      Caption         =   "Order Status => "
      Height          =   360
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3600
      Width           =   3585
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin VB.Frame fraOrderTotals 
      Caption         =   "Order Totals"
      Height          =   2775
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   10227
      Begin VB.TextBox txtTotalIncVat 
         Height          =   288
         Left            =   2220
         TabIndex        =   39
         Top             =   2280
         Width           =   1092
      End
      Begin VB.CommandButton cmdAddUnderpay 
         Caption         =   "&Add Underpay"
         Height          =   360
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2235
         Width           =   1305
      End
      Begin VB.Frame fraPayments 
         BorderStyle     =   0  'None
         Height          =   440
         Left            =   240
         TabIndex        =   30
         Top             =   250
         Width           =   9495
         Begin VB.TextBox txtPayment2 
            Height          =   288
            Left            =   4560
            TabIndex        =   6
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtPayment1 
            Height          =   288
            Left            =   1680
            TabIndex        =   5
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Payment"
            Height          =   255
            Left            =   6840
            TabIndex        =   34
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblTotalPayment 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   8040
            TabIndex        =   33
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblPayment2Caption 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Payment 2"
            Height          =   255
            Left            =   2895
            TabIndex        =   32
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label lblPayment1Caption 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Payment 1"
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame fraMainTotals 
         BorderStyle     =   0  'None
         Height          =   1080
         Left            =   550
         TabIndex        =   25
         Top             =   960
         Width           =   4095
         Begin VB.CommandButton cmdPostageToggle 
            Caption         =   "Inc Postage"
            Height          =   340
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   680
            Width           =   975
         End
         Begin VB.TextBox txtVat 
            Height          =   288
            Left            =   1680
            TabIndex        =   7
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox txtPostage 
            Height          =   288
            Left            =   1680
            TabIndex        =   8
            Top             =   720
            Width           =   1092
         End
         Begin VB.Label lblOrderTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Product Total Inc vat"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "VAT"
            Height          =   255
            Left            =   840
            TabIndex        =   27
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Postage"
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame fraMiscTotals 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   6324
         TabIndex        =   21
         Top             =   960
         Width           =   2655
         Begin VB.TextBox txtReconcilliation 
            Height          =   288
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   1092
         End
         Begin VB.TextBox txtUnderPayment 
            Height          =   288
            Left            =   1440
            TabIndex        =   12
            Top             =   720
            Width           =   1092
         End
         Begin VB.TextBox txtDonation 
            Height          =   288
            Left            =   1440
            TabIndex        =   10
            Top             =   0
            Width           =   1092
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Reconciliation"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Under Payment"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblDonation 
            Alignment       =   1  'Right Justify
            Caption         =   "Donation"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Shape shpUnderPay 
         Height          =   495
         Left            =   120
         Top             =   2160
         Width           =   9975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Inc VAT"
         Height          =   255
         Left            =   1080
         TabIndex        =   41
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblCashbookUnderpayment 
         BackStyle       =   0  'Transparent
         Caption         =   "Cashbook Info"
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   3600
         TabIndex        =   40
         Top             =   2160
         Width           =   4935
      End
      Begin VB.Shape shpMiscTotals 
         Height          =   1215
         Left            =   5174
         Top             =   840
         Width           =   4954
      End
      Begin VB.Shape shpOrderPayments 
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   9987
      End
      Begin VB.Shape shpMainTotals 
         Height          =   1215
         Left            =   120
         Top             =   840
         Width           =   4954
      End
   End
   Begin VB.CommandButton cmdCashBook 
      Caption         =   "&Cash Book"
      Height          =   360
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdNotePad 
      Caption         =   "&Notepad"
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort Order"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   2760
      Top             =   3720
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh Grid"
      Height          =   360
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   1305
   End
   Begin VB.CommandButton cmdInternalNote 
      Caption         =   "&Internal Note"
      Height          =   240
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdConsignNote 
      Caption         =   "&Consignment Note"
      Height          =   360
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1543
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "C&omplete"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   1305
   End
   Begin VB.CommandButton cmdAddProduct 
      Caption         =   "&Add Products"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1305
   End
   Begin MSDBGrid.DBGrid dbgOrderLines 
      Bindings        =   "Order.frx":0000
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "Order.frx":001C
      TabIndex        =   1
      Top             =   1200
      Width           =   10245
   End
   Begin VB.Data datOrderLines 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1092
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   36
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrRoute As String
Dim lbooUsedUpdateTotal As Boolean
Dim lcurTotalCustomerUnderpayment As Currency

Const lconstrReconsolidateMsg = _
            "You MUST now go into the 'cash book' and delete " & vbCrLf & _
            "the underpayment entries which are now no longer " & vbCrLf & _
            "required.  You may also wish to enter a short " & vbCrLf & _
            "consignment note message to inform the customer!"
            
Const lconstrAfterCahsbookMsg = _
            "You may also wish to enter a short " & vbCrLf & _
            "consignment note message to inform the customer!"
Dim lstrScreenHelpFile As String
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property

Private Sub cmdAbort_Click()
Dim lintRetVal As Variant
Dim lstrRetAbort As String
Dim lstrInUseByFlag As String

    frmChildAbortOptions.Style = 0
    frmChildAbortOptions.Show vbModal
    lstrRetAbort = frmChildAbortOptions.AbortOption
    Unload frmChildAbortOptions
    
    Select Case lstrRetAbort
    Case "BACK"
        Exit Sub
    Case "SAVE"
        'Will never see Abort button unless order entry
        lintRetVal = MsgBox("Do you wish to Save this order so far and set to cancelled?", vbYesNo + vbExclamation, gconstrTitlPrefix & "Abort Order")
        If lintRetVal = vbYes Then
            SaveLocalFields
            lstrInUseByFlag = LockingPhaseGen(True)
            AddAdviceNote lstrInUseByFlag, "X"
            GetAdviceOrderNum lstrInUseByFlag, gstrAdviceNoteOrder.lngCustNum
                
            MsgBox "Your Order number is " & gstrAdviceNoteOrder.lngOrderNum, , gconstrTitlPrefix & "Order Incomplete!"
            
            UpdateSalesCode
        
            AppendOrderLinesToMaster gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum
            UpdateMasterStockRecords gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum, Now()
            With gstrAdviceNoteOrder
                AddNewFileHistoryItem .lngCustNum, .lngOrderNum, _
                    Trim$(Trim$(gstrCustomerAccount.strSalutation) & " " & _
                    Trim$(gstrCustomerAccount.strInitials) & " " & Trim$(gstrCustomerAccount.strSurname)), _
                    "Aborted"
            End With
            
            ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
            ToggleAdviceInUseBy gstrAdviceNoteOrder.lngOrderNum, False
        Else
            Exit Sub
        End If
    Case "ABORT"
        lintRetVal = MsgBox("WARNING: By aborting this order no information will be saved!", vbYesNo + vbExclamation, gconstrTitlPrefix & "Abort Order")
        If lintRetVal = vbYes Then
            ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
        Else
            Exit Sub
        End If
    End Select
    
    ClearCustomerAcount
    ClearAdviceNote
    ClearGen
    
    gintOrderLineNumber = 0
    
    Unload Me
    frmAbout.Show
    
End Sub
Private Sub cmdAddProduct_Click()

    Busy True, Me
    Load frmChildProducts
    Busy False, Me
    frmChildProducts.Show vbModal
    Busy True, Me
    UpdateOrdLinesWithProds gstrCustomerAccount.lngCustNum
    UpdateOrderLinesTotals
    datOrderLines.Refresh
    dbgOrderLines.Refresh
    Busy False, Me
    On Error GoTo 0
    dbgOrderLines.SetFocus
    On Error Resume Next
        
End Sub
Private Sub cmdAddUnderpay_Click()
Dim lstrUndepayValue As String
Dim lstrSQL As String

        lstrUndepayValue = SystemPrice(PriceVal(InputBox("Please enter the Amount to be added to this order, eg. 25.50" & _
            vbCrLf & vbCrLf & "If you have clicked this button in error, or" & _
            vbCrLf & "do not wish to proceed, please enter 0", _
            "Underpayment Amount", SystemPrice(CStr(lcurTotalCustomerUnderpayment)))))
        
        If PriceVal(lstrUndepayValue) = 0 Then
            MsgBox "You must enter a valid value", , gconstrTitlPrefix & "Add Underpayment"
           
            Exit Sub
        End If
        
       
        lstrSQL = "INSERT INTO " & gtblOrderLines & " ( CatNum, ItemDescription, Qty, Price, Custnum, TaxCode, Weight, OrderLineNum, BinLocation, TotalWeight ) " & _
            "SELECT 'UNDERPAY' AS Expr1, 'Underpayment from previous order' AS Expr2, " & _
            " 1 AS Expr4, " & CCur(lstrUndepayValue) & " AS Expr5, " & _
            gstrCustomerAccount.lngCustNum & " as Expr6, 'Z' as Expr7, " & _
            "0 as Expr8, 999 as LineNum, '" & OneSpace(" ") & "' AS Expr9, '" & OneSpace(" ") & "' AS Expr10;"


        gdatLocalDatabase.Execute lstrSQL
        UpdateOrderLinesTotals
        datOrderLines.Refresh
        dbgOrderLines.Refresh
        On Error GoTo 0
        dbgOrderLines.SetFocus
        On Error Resume Next
        
        MsgBox lconstrReconsolidateMsg, vbInformation, gconstrTitlPrefix & "Add Underpayment"
        
        UpdateTotal
        
End Sub

Private Sub cmdBack_Click()

    SaveLocalFields
    Set gstrCurrentLoadedForm = frmOrdDetails
    frmOrdDetails.Route = Me.Route
    frmOrdDetails.Show
    Unload Me
        
End Sub

Private Sub cmdCashBook_Click()

    Busy True, Me
    frmChildCashbook.Route = gconstrCashbookSpecificCustomer
    Busy False, Me
    frmChildCashbook.Show vbModal
    MsgBox lconstrAfterCahsbookMsg, vbInformation, gconstrTitlPrefix & "Customer Cashbook"
        
End Sub

Private Sub cmdConsignNote_Click()
Dim lstrLockingFlag As String

    lstrLockingFlag = LockingPhaseGen(True)
    
    gstrConsignmentNote.strType = "Consignment"

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
        ToggleRemarkInUseBy gstrConsignmentNote.lngRemarkNumber, False
        frmChildNote.NoteText = gstrConsignmentNote.strText
        frmChildNote.NoteType = "Consignment Note Comments"
        Load frmChildNote
        frmChildNote.Show vbModal
    End If
    
    gstrConsignmentNote.strText = frmChildNote.NoteText
    gstrConsignmentNote.strType = frmChildNote.NoteType
    UpdateRemark gstrAdviceNoteOrder.lngConsignRemarkNum, gstrConsignmentNote.strType, gstrConsignmentNote.strText

End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdInternalNote_Click()
Dim lstrLockingFlag As String

    lstrLockingFlag = Trim$(gstrGenSysInfo.strUserName) & " " & Now()
    gstrInternalNote.strType = "Internal"

    If gstrAdviceNoteOrder.lngAdviceRemarkNum <> 0 Then
        If gstrInternalNote.lngRemarkNumber <> 0 Then
            frmChildNote.NoteText = gstrInternalNote.strText
            frmChildNote.NoteType = "Internal Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        Else
            gstrInternalNote.lngRemarkNumber = gstrAdviceNoteOrder.lngAdviceRemarkNum
            GetRemark gstrInternalNote.lngRemarkNumber, gstrInternalNote
            gstrAdviceNoteOrder.lngAdviceRemarkNum = gstrInternalNote.lngRemarkNumber
            frmChildNote.NoteText = gstrInternalNote.strText
            frmChildNote.NoteType = "Internal Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
            
            
        End If
    Else
        AddNewRemark lstrLockingFlag
        GetRemarkNum lstrLockingFlag, gstrInternalNote
        gstrAdviceNoteOrder.lngAdviceRemarkNum = gstrInternalNote.lngRemarkNumber
        frmChildNote.NoteText = gstrInternalNote.strText
        frmChildNote.NoteType = "Internal Note Comments"
        Load frmChildNote
        frmChildNote.Show vbModal
    End If
    
    gstrInternalNote.strText = frmChildNote.NoteText
    gstrInternalNote.strType = frmChildNote.NoteType
    UpdateRemark gstrAdviceNoteOrder.lngAdviceRemarkNum, gstrInternalNote.strType, gstrInternalNote.strText

End Sub

Private Sub cmdNext_Click()
Dim lstrInUseByFlag As String
Dim lstrFileName As String
Dim lintRetVal As Variant
Dim lstrRecordOrderStatus As String

    UpdateTotal
    
    If UCase$(Mid$(lblPayment1Caption, 10, 6)) = "CHEQUE" Then
        If CCur(txtPayment1) = 0 Then
            MsgBox "You have not enter a value for a cheque payment." & vbCrLf & _
                "Please go back and do so!", , gconstrTitlPrefix & "Mandatory Field"
            If txtPayment1.Visible = True Then
                txtPayment1.SetFocus
            End If
            Exit Sub
        End If
    End If
    
    If UCase$(Mid$(lblPayment2Caption, 10, 6)) = "CHEQUE" Then
        If CCur(txtPayment2) = 0 Then
            MsgBox "You have not enter a value for a cheque payment." & vbCrLf & _
                "Please go back and do so!", , gconstrTitlPrefix & "Mandatory Field"
            If txtPayment2.Visible = True Then
                txtPayment2.SetFocus
            End If
            Exit Sub
        End If
    End If
    
    SaveLocalFields

    lstrInUseByFlag = LockingPhaseGen(True) 
    
    ShowStatus 12
    If Me.Route <> gconstrOrderModify Then
        AddAdviceNote lstrInUseByFlag, gstrOrderEntryOrderStatus
        GetAdviceOrderNum lstrInUseByFlag, gstrAdviceNoteOrder.lngCustNum
            
        MsgBox "Your Order number is " & gstrAdviceNoteOrder.lngOrderNum, , gconstrTitlPrefix & "Order Complete!"
        
        UpdateSalesCode
    
        AppendOrderLinesToMaster gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum
        UpdateMasterStockRecords gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum, Now()
        With gstrAdviceNoteOrder
            AddNewFileHistoryItem .lngCustNum, .lngOrderNum, _
                Trim$(Trim$(gstrCustomerAccount.strSalutation) & " " & _
                Trim$(gstrCustomerAccount.strInitials) & " " & Trim$(gstrCustomerAccount.strSurname)), _
                "Ordered"
        End With
    Else
        lstrRecordOrderStatus = GetAdviceOrderStatus(gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum)
        If lstrRecordOrderStatus = "X" Then
            lintRetVal = MsgBox("The status of this order is currently set to cancelled! Do you wish" & vbCrLf & _
                "to set it to 'A' (Awaiting Packing), so it can be processed?", vbYesNo, gconstrTitlPrefix & "Order Status!")
            If lintRetVal = vbYes Then
                UpdateOrderStatus "A", 0, "S", gstrAdviceNoteOrder.lngOrderNum
            End If
        End If
        If Trim$(gstrOrderEntryOrderStatus) <> "" Then
            If lstrRecordOrderStatus <> gstrOrderEntryOrderStatus Then
                UpdateOrderStatus gstrOrderEntryOrderStatus, 0, "S", gstrAdviceNoteOrder.lngOrderNum
            End If
        End If
        UpdateAdviceNote
        AppendNewOrderLinesToMaster gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum
        UpdateOrderLinesToMaster gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum
        With gstrAdviceNoteOrder
            AddNewFileHistoryItem .lngCustNum, .lngOrderNum, _
                Trim$(Trim$(.strSalutation) & " " & _
                Trim$(.strInitials) & " " & Trim$(.strSurname)), _
                "Order History"
        End With
    End If
        
    ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
    ToggleAdviceInUseBy gstrAdviceNoteOrder.lngOrderNum, False
    
    If UCase$(Mid$(lblPayment1Caption, 10, 11)) = "CREDIT CARD" Then
        If PriceVal(txtPayment1) > PriceVal(gstrReferenceInfo.strPostageWaiveratio) Then
            'set Status to Hold and await Authorisation!
            lintRetVal = MsgBox("The amount to be charged to the Credit Card is Over " & gstrReferenceInfo.strPostageWaiveratio & "," & vbCrLf & _
                "Would you like to hold this order for Authorisation?", vbYesNo + vbDefaultButton1, gconstrTitlPrefix & "Authorisation Required")
            If lintRetVal = vbYes Then
                UpdateOrderStatus "H", 0, "S", gstrAdviceNoteOrder.lngOrderNum
            End If
        End If
    ElseIf UCase$(Mid$(lblPayment2Caption, 10, 11)) = "CREDIT CARD" Then
        If PriceVal(txtPayment2) > PriceVal(gstrReferenceInfo.strPostageWaiveratio) Then
            'set Status to Hold and await Authorisation!
            lintRetVal = MsgBox("The amount to be charged to the Credit Card is Over " & gstrReferenceInfo.strPostageWaiveratio & "," & vbCrLf & _
                "Would you like to hold this order for Authorisation?", vbYesNo + vbDefaultButton1, gconstrTitlPrefix & "Authorisation Required")
            If lintRetVal = vbYes Then
                UpdateOrderStatus "H", 0, "S", gstrAdviceNoteOrder.lngOrderNum
            End If
        End If
    End If
    
    ClearCustomerAcount
    ClearAdviceNote
    ClearGen
    
    gintOrderLineNumber = 0
    
    Unload Me
    frmAbout.Show
    
End Sub
Sub SaveLocalFields()
    
    With gstrAdviceNoteOrder
        .strDonation = SystemPrice(txtDonation)
        .strPayment = SystemPrice(txtPayment1)
        .strPayment2 = SystemPrice(txtPayment2)
                
        .strPostage = SystemPrice(txtPostage)
        .strVAT = SystemPrice(txtVat)
        
        .strTotalIncVat = SystemPrice(txtTotalIncVat)
        .strUnderpayment = SystemPrice(txtUnderPayment)
        .strReconcilliation = SystemPrice(txtReconcilliation)
    End With

End Sub
Sub GetLocalFields()

    With gstrAdviceNoteOrder
        txtDonation = SystemPrice(.strDonation)
        txtPayment1 = SystemPrice(.strPayment)
        txtPayment2 = SystemPrice(.strPayment2)
        txtPostage = SystemPrice(.strPostage)
        txtVat = SystemPrice(.strVAT)
        
        txtTotalIncVat = SystemPrice(.strTotalIncVat)
        txtUnderPayment = SystemPrice(.strUnderpayment)
        txtReconcilliation = SystemPrice(.strReconcilliation)
    End With
    
End Sub

Private Sub cmdNotePad_Click()

    frmChildCuNotes.Show vbModal
    
End Sub

Private Sub cmdOrderStatus_Click()

    With frmChildGenericDropdown
        .LabelStr = ""
        .FormCaption = "Wine Order Selection"
        .LabelCaption = ""
        .CodeField = "ListCode"
        .DescField = "Description"
        .AddStar = False
        .DB = "LOCAL"
        .SQL = "SELECT " & gtblLists & ".ListName, " & gtblListDetails & ".ListCode, " & gtblListDetails & ".Description " & _
            "FROM " & gtblLists & " INNER JOIN " & gtblListDetails & " ON " & gtblLists & ".ListNum = " & gtblListDetails & ".ListNum " & _
            "WHERE (((" & gtblLists & ".ListName)='Order Status') AND ((" & gtblListDetails & ".ListCode)='A' Or (" & gtblListDetails & ".ListCode)='W'));"
        Debug.Print .SQL
        .Show vbModal
        If .Cancelled = True Then
            Exit Sub
        End If
        gstrOrderEntryOrderStatus = .ReturnCode
    End With
    
    cmdOrderStatus.Caption = "Order Status => " & gstrOrderEntryOrderStatus
    
End Sub

Private Sub cmdPostageToggle_Click()

    Select Case cmdPostageToggle.Caption
    Case "Inc Postage"
        cmdPostageToggle.Caption = "Ex Postage"
    Case "Ex Postage"
        cmdPostageToggle.Caption = "Alt Postage"
    Case "Alt Postage"
        cmdPostageToggle.Caption = "Inc Postage"
    End Select

    UpdateTotal
    
End Sub
Private Sub cmdRefresh_Click()
    
    UpdateOrderLinesTotals
    datOrderLines.Refresh
    dbgOrderLines.Refresh
    
End Sub
Sub UpdateTotal()
Dim lintOrderSubTotal As Integer
Dim lbooPostageSetHere As Boolean
Dim lstrPostageCode As String
Dim lstrPostageListVars As ListVars

    On Error GoTo ErrHandler
    
    lbooUsedUpdateTotal = True
    lbooPostageSetHere = False
    OrderTotal gstrAdviceNoteOrder.lngCustNum
    txtVat = SystemPrice(gstrVatTotal)
    lblOrderTotal = SystemPrice(gstrOrderTotal)
    
    txtTotalIncVat = SystemPrice(CCur(lblOrderTotal) + CCur(txtPostage) + _
        CCur(txtVat) + CCur(txtDonation))
    
    If CCur(lblOrderTotal) < CCur(gstrReferenceInfo.strPostageWaiveratio) Then
    
    Else
        If cmdPostageToggle.Caption <> "Alt Postage" Then
            txtPostage = SystemPrice(CCur(0))
        End If
        
        lbooPostageSetHere = True
    End If
    
    Select Case cmdPostageToggle.Caption
    Case "Inc Postage"
        If lbooPostageSetHere = False Then
            lstrPostageListVars.strListName = "PForce Service Indicator"
            
            Select Case UCase$(Trim$(gstrAdviceNoteOrder.strCourierCode))
            Case "PF 48", "", "PF" 
                lstrPostageCode = "SUP" 
            Case Else 
                lstrPostageCode = Trim$(gstrAdviceNoteOrder.strCourierCode) 
            End Select
        
            lstrPostageListVars.strListCode = lstrPostageCode 
            GetListVarsAll lstrPostageListVars
            
            If Trim$(lstrPostageListVars.strUserDef2) = "" Then
                MsgBox "Please contact your IT Support office and ask them to set your postage prices!", vbInformation, gconstrTitlPrefix & "Update Total"
                lstrPostageListVars.strUserDef2 = "0"
            End If
            txtPostage = SystemPrice(CCur(lstrPostageListVars.strUserDef2))
        End If
    Case "Ex Postage"
        txtPostage = SystemPrice(CCur(0))
    Case "Alt Postage"
        txtPostage = SystemPrice(CCur(txtPostage))
    End Select
    
    txtTotalIncVat = SystemPrice(CCur(lblOrderTotal) + CCur(txtPostage) + _
        CCur(txtDonation))
    
    If UCase$(Mid$(lblPayment1Caption, 10, 11)) = "CREDIT CARD" Then
        txtPayment1 = SystemPrice(CCur(txtTotalIncVat) - CCur(txtPayment2))
    ElseIf UCase$(Mid$(lblPayment2Caption, 10, 11)) = "CREDIT CARD" Then
        txtPayment2 = SystemPrice(CCur(txtTotalIncVat) - CCur(txtPayment1))
    End If
    
    lblTotalPayment = SystemPrice(CCur(txtPayment1) + CCur(txtPayment2))
    
        If CCur(lblTotalPayment) > CCur(txtTotalIncVat) Then
            txtUnderPayment = SystemPrice("0")
        Else
            txtUnderPayment = SystemPrice(CCur(txtTotalIncVat) - CCur(lblTotalPayment))
        End If
        If CCur(lblTotalPayment) > CCur(txtTotalIncVat) Then
            txtReconcilliation = SystemPrice(CCur(lblTotalPayment) - CCur(txtTotalIncVat))
        Else
            txtReconcilliation = SystemPrice("0")
        End If
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmOrder.UpdateTotal", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub

Private Sub dbgOrderLines_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    UpdateOrderLinesTotals
    UpdateTotal
    dbgOrderLines.Refresh

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    If gbooFreeOrder = True Then
        MsgBox "As the order has no charge, remember to changes prices for each product to Zero!", , gconstrTitlPrefix & "Free Order"
    End If
    
    lbooUsedUpdateTotal = False
    
    With gstrAdviceNoteOrder
        lblPayment1Caption = "Payment (" & GetListCodeDesc("Payment Method", .strPaymentType1) & ")"
        If lblPayment1Caption = "Payment ()" Then
            lblPayment1Caption.Visible = False
            txtPayment1.Visible = False
        Else
            lblPayment1Caption.Visible = True
            txtPayment1.Visible = True
        End If
        
        lblPayment2Caption = "Payment (" & GetListCodeDesc("Payment Method", .strPaymentType2) & ")"
        If lblPayment2Caption = "Payment ()" Then
            lblPayment2Caption.Visible = False
            txtPayment2.Visible = False
        Else
            lblPayment2Caption.Visible = True
            txtPayment2.Visible = True
        End If
    End With
            
    If gbooFreeOrder = True Then
        lblPayment2Caption.Visible = False
        txtPayment2.Visible = False
        lblPayment1Caption.Visible = False
        txtPayment1.Visible = False
        cmdPostageToggle.Caption = "Ex Postage"
    End If
    
    Select Case Me.Route
    Case gconstrAccount
        cmdAbort.Visible = False
        cmdOrderStatus.Visible = False
    Case gconstrEntry
        cmdAbort.Visible = True
        cmdOrderStatus.Visible = True
        If gstrOrderEntryOrderStatus = " " Then
            gstrOrderEntryOrderStatus = "A"
        End If
        cmdOrderStatus.Caption = "Order Status =>  " & AmpersandDouble(GetListCodeDesc("Order Status", gstrOrderEntryOrderStatus))
    Case gconstrOrderModify, gconstrEnquiry
        cmdAbort.Visible = False
        cmdOrderStatus.Visible = True
    End Select
            
    ShowBanner Me, Me.Route
    
    GetLocalFields
    
    On Error GoTo ErrHandler
    
    If Me.Route = gconstrOrderModify Then
        ApndOrdLinesMstToLocal gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum
    End If
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datOrderLines.DatabaseName = gstrStatic.strLocalTestingDBFile
    Case gconstrLiveMode
        datOrderLines.DatabaseName = gstrStatic.strLocalDBFile
    End Select
    
   
    If gstrSystemRoute <> srCompanyRoute Then
        datOrderLines.Connect = gstrDBPasswords.strLocalDBPasswordString
    End If
    
    If gstrSystemRoute = srStandardRoute Then
        cmdOrderStatus.Visible = False
    End If
    
    'Converted table names to constants
    datOrderLines.RecordSource = "select * from " & gtblOrderLines & " where CustNum = " & _
        gstrAdviceNoteOrder.lngCustNum & " order by OrderLineNum, ItemDescription"
    datOrderLines.Refresh

    lcurTotalCustomerUnderpayment = CCur(AccountBalance(gstrAdviceNoteOrder.lngCustNum))
    If Val(lcurTotalCustomerUnderpayment) > 0 Then
        lblCashbookUnderpayment = "This customer owes " & SystemPrice(CStr(lcurTotalCustomerUnderpayment)) & " from previous orders!"
        lblCashbookUnderpayment.ForeColor = vbRed
        cmdAddUnderpay.Enabled = True
    Else
        lblCashbookUnderpayment = "This customer owes nothing from previous orders! " & vbCrLf & "(Please check the cashbook!!)"
        lblCashbookUnderpayment.ForeColor = vbButtonText
        cmdAddUnderpay.Enabled = False
    End If

    If gstrReferenceInfo.booDonationAvail = False Then
        lblDonation.Visible = False
        txtDonation.Visible = False
    End If
   
    UpdateTotal
    
    SetupHelpFileReqs
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmOrder.Load", "Local")
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
Dim llngFormHalfWidth As Long

    On Error Resume Next
    
    llngFormHalfWidth = Me.Width / 2
    
    With cmdNext
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdNext.Left - (cmdBack.Width + 120)
    End With

    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With
    
    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
    With cmdAbort
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelp.Width + cmdHelp.Left + 120
    End With
    
    With cmdNotePad
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdAbort.Width + cmdAbort.Left + 120
    End With
    
    With cmdCashBook
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdNotePad.Width + cmdNotePad.Left + 120
    End With
    
    With fraOrderTotals
        .Left = 189
        .Width = Me.Width - (189 * 2)
        .Top = (cmdHelp.Top - 305) - fraOrderTotals.Height
    End With
    
    With cmdAddProduct
        .Left = cmdHelp.Left
        .Top = (fraOrderTotals.Top - 120) - .Height
    End With
    
    With cmdRefresh
        .Left = cmdAbort.Left
        .Top = cmdAddProduct.Top
    End With
    
    With cmdConsignNote
        .Left = cmdNext.Left - 240
        .Top = cmdAddProduct.Top
    End With
    
    With cmdInternalNote
        .Left = cmdBack.Left - 480
        .Top = cmdAddProduct.Top
    End With
    
    With cmdOrderStatus
        .Left = (cmdConsignNote.Left - .Width) - 240
        .Top = cmdAddProduct.Top
    End With
    
    With dbgOrderLines
        .Width = Me.Width - 360
        If .Top > cmdAddProduct.Top Then
            .Height = (.Top - cmdAddProduct.Top) - 120
        
        Else
            .Height = (cmdAddProduct.Top - .Top) - 120
        End If
    End With
    
    With shpOrderPayments
        If fraOrderTotals.Width > (120 * 2) Then
            .Width = fraOrderTotals.Width - (120 * 2)
        Else
            .Width = (120 * 2) - fraOrderTotals.Width
        End If
    End With
    
    With fraPayments
        .Left = (fraOrderTotals.Width / 2) - (.Width / 2)
    End With
    
    With shpMiscTotals
        .Left = (fraOrderTotals.Width / 2) + 60
        If (fraOrderTotals.Width / 2) > 160 Then
            .Width = (fraOrderTotals.Width / 2) - 160
        Else
            .Width = 160 - (fraOrderTotals.Width / 2)
        End If
    End With
    
    With fraMiscTotals
        .Left = shpMiscTotals.Left + (shpMiscTotals.Width / 2) - (.Width / 2)
    End With
    
    With shpMainTotals
        If (fraOrderTotals.Width / 2) > 160 Then
            .Width = (fraOrderTotals.Width / 2) - 160
        Else
            .Width = 160 - (fraOrderTotals.Width / 2)
        End If
    End With
    
    With fraMainTotals
        .Left = shpMainTotals.Left + (shpMainTotals.Width / 2) - (.Width / 2)
    End With
    
    With shpUnderPay
        If fraOrderTotals.Width > 1922 Then
            .Width = (fraOrderTotals.Width - 1922)
        Else
            .Width = 1922 - fraOrderTotals.Width
        End If
    End With
    
    With shpUnderPay
        .Left = shpOrderPayments.Left
        .Width = shpOrderPayments.Width
    End With
    
End Sub
Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
Private Sub txtDonation_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtDonation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtReconcilliation.SetFocus
    End If
    
End Sub

Private Sub txtDonation_LostFocus()

    txtDonation = SystemPrice(txtDonation)
    UpdateTotal
    
End Sub

Private Sub txtPayment1_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = 13 Then 'Carriage return
        If txtPayment2.Visible Then
            txtPayment2.SetFocus
        Else
            txtVat.SetFocus
        End If
    End If
    
End Sub

Private Sub txtPayment2_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = 13 Then 'Carriage return
        txtVat.SetFocus
    End If
    
End Sub

Private Sub txtPostage_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = 13 Then 'Carriage return
        If txtDonation.Visible = True Then
            txtDonation.SetFocus
        Else
            txtReconcilliation.SetFocus
        End If
    End If
    
End Sub

Private Sub txtReconcilliation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtUnderPayment.SetFocus
    End If
    
End Sub

Private Sub txtTotalIncVat_GotFocus()

    SetSelected Me

End Sub
Private Sub txtPostage_GotFocus()

    SetSelected Me

End Sub

Private Sub txtTotalIncVat_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdNext.SetFocus
    End If
    
End Sub

Private Sub txtUnderPayment_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtTotalIncVat.SetFocus
    End If
    
End Sub

Private Sub txtVat_GotFocus()

    SetSelected Me

End Sub
Private Sub txtPayment2_GotFocus()

    SetSelected Me

End Sub
Private Sub txtPayment1_GotFocus()

    SetSelected Me

End Sub
Private Sub txtPostage_LostFocus()

    txtPostage = SystemPrice(txtPostage)
    UpdateTotal
    
End Sub
Private Sub txtTotalIncVat_LostFocus()

    txtTotalIncVat = SystemPrice(txtTotalIncVat)
    UpdateTotal
    
End Sub

Private Sub txtVat_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = 13 Then 'Carriage return
        txtPostage.SetFocus
    End If
    
End Sub

Private Sub txtVat_LostFocus()

    txtVat = SystemPrice(txtVat)
    UpdateTotal
    
End Sub
Private Sub txtPayment2_LostFocus()

    txtPayment2 = SystemPrice(txtPayment2)
    UpdateTotal
End Sub
Private Sub txtPayment1_LostFocus()

    txtPayment1 = SystemPrice(txtPayment1)
    UpdateTotal
    
End Sub
Private Sub txtReconcilliation_GotFocus()

    SetSelected Me

End Sub
Private Sub txtReconcilliation_LostFocus()

    txtReconcilliation = SystemPrice(txtReconcilliation)
    UpdateTotal
    
End Sub
Private Sub txtUnderPayment_GotFocus()

    SetSelected Me

End Sub
Private Sub txtUnderPayment_LostFocus()

    txtUnderPayment = SystemPrice(txtUnderPayment)
    UpdateTotal
    
End Sub
Sub SetupHelpFileReqs()
    
    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/Order.xml>WhatsScreen"
    
    ctlBanner1.WhatsThisHelpID = IDH_ORDER_MAIN
    ctlBanner1.WhatIsID = IDH_ORDER_MAIN
    
    ctlBottomLine1.WhatsThisHelpID = IDH_ORDER_MAIN
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    cmdNext.WhatsThisHelpID = IDH_STANDARD_NEXT
    cmdAbort.WhatsThisHelpID = IDH_STANDARD_ABORT
    cmdNotePad.WhatsThisHelpID = IDH_STANDARD_CUNOTES
    cmdCashBook.WhatsThisHelpID = IDH_STANDARD_CHICASHBOOK
    cmdConsignNote.WhatsThisHelpID = IDH_STANDARD_CONSNOTE
    
    cmdAddProduct.WhatsThisHelpID = IDH_ORDER_ADDPRODS
    cmdRefresh.WhatsThisHelpID = IDH_ORDER_REFRESH
    cmdPostageToggle.WhatsThisHelpID = IDH_ORDER_POSTAGE
    cmdAddUnderpay.WhatsThisHelpID = IDH_ORDER_UNDERPAY

    dbgOrderLines.WhatsThisHelpID = IDH_ORDER_GRIDOL
    txtPayment1.WhatsThisHelpID = IDH_ORDER_PAY1
    txtPayment2.WhatsThisHelpID = IDH_ORDER_PAY2
    lblTotalPayment.WhatsThisHelpID = IDH_ORDER_LBTOTPAY
    txtVat.WhatsThisHelpID = IDH_ORDER_VAT
    lblOrderTotal.WhatsThisHelpID = IDH_ORDER_LBLORDTOT
    txtPostage.WhatsThisHelpID = IDH_ORDER_TXTPOST
    txtTotalIncVat.WhatsThisHelpID = IDH_ORDER_TOTINCVAT
    txtDonation.WhatsThisHelpID = IDH_ORDER_DONAT
    txtReconcilliation.WhatsThisHelpID = IDH_ORDER_RECON
    txtUnderPayment.WhatsThisHelpID = IDH_ORDER_TXTUNDERP
    lblCashbookUnderpayment.WhatsThisHelpID = IDH_ORDER_LBLCASH

End Sub
