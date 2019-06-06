VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCheque 
   Caption         =   "Please Select a Acount..."
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
   Begin VB.CheckBox chkGenericChk 
      Caption         =   "Check1"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdViewAdviceNote 
      Caption         =   "&View Advice Note"
      Height          =   360
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1543
   End
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7235
      Width           =   1305
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
      Left            =   6600
      TabIndex        =   13
      Top             =   1200
      Width           =   3855
      Begin VB.OptionButton optSearchField 
         Caption         =   "&Customer Number"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   "CustNumber"
         Top             =   480
         Value           =   -1  'True
         Width           =   1692
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "&Order Number"
         Height          =   252
         Index           =   2
         Left            =   1920
         TabIndex        =   5
         Tag             =   "OrderNum"
         Top             =   240
         Width           =   1452
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "Customer &Name"
         Height          =   252
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Tag             =   "Name"
         Top             =   480
         Width           =   1812
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "C&heque Number"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Tag             =   "BPCS"
         Top             =   240
         Width           =   1812
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Data datCashbook 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   1140
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
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   8400
      Top             =   840
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   17
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin MSDBGrid.DBGrid dbgCashbook 
      Bindings        =   "Cheque.frx":0000
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "Cheque.frx":001A
      TabIndex        =   9
      Top             =   2760
      Width           =   10095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Order(s) Credit Card payments"
            Key             =   "OCC"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Order(s) Under payments"
            Key             =   "OUP"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Refund(s) Cheque payments"
            Key             =   "RCQ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Refund(s) Credit Card Payments"
            Key             =   "RCC"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   "To remove a cheque entry, enter Bank Rep Print Date of 31/12/1999"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   6720
      Width           =   5295
   End
   Begin VB.Label lblFoundNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Found 0 records"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Criteria :-"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrRoute As String
Dim lstrScreenHelpFile As String
Dim lstrExtraSQL As String
Dim lstrWhereClause As String
Dim lstrLastTabItem As String
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property


Private Sub chkGenericChk_Click()

    RefreshGrid TabStrip1.SelectedItem.Key
    
End Sub

Public Sub cmdBack_Click()

    Select Case Me.Route
    Case gconstrCashbookSpecificCustomer
        Unload Me
    Case Else
        Unload Me
        ClearAdviceNote
        ClearCustomerAcount
        ClearGen
        frmAbout.Show
    End Select
    
End Sub

Private Sub cmdFind_Click()
Const lstrEndOfMessage = ", or select a different search method"
Dim llngOrderNum As Long
Dim llngCustNum As Long
Dim lstrCustName As String

    If optSearchField(0).Value = True Then  ' Cheque Number
        If Trim$(txtSearchCriteria) <> "" Then
            lstrExtraSQL = "where ChequeNum = '" & Trim$(txtSearchCriteria) & "' and "
        End If

    ElseIf optSearchField(1).Value = True Then ' Customer Number
        If CLng(Val(txtSearchCriteria)) = 0 Then
            MsgBox "You must enter a Customer Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where CustNum = " & CLng(Val(txtSearchCriteria)) & " and "
        End If

    ElseIf optSearchField(2).Value = True Then ' Order Number
        If CLng(Val(txtSearchCriteria)) = 0 Then
            MsgBox "You must enter a Order Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where OrderNum = " & CLng(txtSearchCriteria) & " and "
        End If

    ElseIf optSearchField(3).Value = True Then ' Customer name
        If Trim$(txtSearchCriteria) = "" Then
            MsgBox "You must enter a Customer Name" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "Where CardName = '" & Trim$(txtSearchCriteria) & "' and "
        End If

    End If

    Busy True, Me
    
    RefreshGrid TabStrip1.SelectedItem.Key
    
    Busy False, Me

    If datCashbook.Recordset.RecordCount > 0 Then
        On Error Resume Next
        llngCustNum = CLng(datCashbook.Recordset("CustNum"))
        llngOrderNum = CLng(datCashbook.Recordset("OrderNum"))
        lstrCustName = Trim$(datCashbook.Recordset("CardName")) & ""
        AddNewFileHistoryItem llngCustNum, llngOrderNum, lstrCustName, "Cash Book"
    End If

End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdViewAdviceNote_Click()
Dim llngOrderNum As Long
Dim lstrOrderStatus As String

    On Error Resume Next
    llngOrderNum = datCashbook.Recordset("OrderNum")
    On Error GoTo 0
    If llngOrderNum = 0 Then
        MsgBox "This does not appear to be an Order Number, please select from grid!", , gconstrTitlPrefix & "Mandatory Field"
        Exit Sub
    End If
    
    If UCase$(App.ProductName) = "LITE" Then
        ChooseLayout ltAdviceWithAddress, Me
    Else
        ChooseLayout ltAdviceNote, Me
    End If
    
    Busy True, Me
    
    If datCashbook.Recordset("Reconcil") > 0 Then
        PrintObjAdviceNotesGeneral 0, 0, "S", llngOrderNum, , "R"
    Else
        PrintObjAdviceNotesGeneral 0, 0, "S", llngOrderNum
    End If
    
    Busy False, Me
    ShowPlotReport
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    On Error GoTo ErrHandler
    
    If gbooSQLServerInUse = True Then
        datCashbook.Connect = "ODBC;DATABASE=Mmos;DSN=Mmos;"
    
    Else
        Select Case gstrUserMode
        Case gconstrTestingMode
            datCashbook.DatabaseName = gstrStatic.strCentralTestingDBFile
        Case gconstrLiveMode
            datCashbook.DatabaseName = gstrStatic.strCentralDBFile
        End Select
        
        If gstrSystemRoute <> srCompanyRoute Then
            datCashbook.Connect = gstrDBPasswords.strCentralDBPasswordString
        End If
    
    End If
    
    Select Case Me.Route
    Case gconstrCashbookSpecificCustomer
        MsgBox "Error please report where this happened!", vbInformation, gconstrTitlPrefix & "Child Cash Book"
    Case Else
        dbgCashbook.AllowDelete = False
        lblDelete.Visible = False
    End Select
        
    lstrExtraSQL = "WHERE 1=0;"
    
    RefreshGrid "OCC"
    
    ShowBanner Me, Me.Route

    GetLocalFields
    
    SetupHelpFileReqs

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmCheque.Load", "Central", True)
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
    
    With lblFoundNumber
        .Top = (cmdHelp.Top - .Height) - 305
    End With
    
    With lblDelete
        .Top = (cmdHelp.Top - .Height) - 305
    End With
    
    With TabStrip1
        .Width = Me.Width - 360
        If (cmdHelp.Top - .Top) > 665 Then
            .Height = (cmdHelp.Top - .Top) - 665
        Else
            .Height = 665 - (.Top - cmdHelp.Top)
        End If
    End With
    
    With dbgCashbook
        .Width = TabStrip1.Width - 360 '360 '240
        .Height = TabStrip1.Height - 800
    End With
        
End Sub

Private Sub TabStrip1_Click()

    RefreshGrid TabStrip1.SelectedItem.Key
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
        
End Sub
Sub GetLocalFields()

End Sub
Sub SaveLocalFields()

End Sub

Private Sub txtSearchCriteria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdFind_Click
    End If
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/Finance.xml>WhatsScreen"
    
    ctlBanner1.WhatsThisHelpID = IDH_FINANCE_MAIN
    ctlBanner1.WhatIsID = IDH_FINANCE_MAIN
    
    ctlBottomLine1.WhatsThisHelpID = IDH_FINANCE_MAIN
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    cmdFind.WhatsThisHelpID = IDH_STANDARD_FIND
    lblFoundNumber.WhatsThisHelpID = IDH_STANDARD_LBLGRITOTFOUND
    
    fraSearchBy.WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(0).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(1).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(2).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(3).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    
    txtSearchCriteria.WhatsThisHelpID = IDH_FINANCE_SEARCHCRIT
    
End Sub
Sub RefreshGrid(pstrKey As String)

Const lconOrdCCSQL = "SELECT CustNum, OrderNum, Trim(Denom) & Format(IIf(Trim$([OrderType])='C',[Payment],0)" & _
    "+IIf(Trim$([PaymentType2])='C',[Payment2],0),'0.00') AS Amount, BankRepPrintDate, CardNumber, CardName, CardStartDate, " & _
    "CardIssueNumber as IssueNum, CardType, AuthorisationCode as AuthCode FROM " & gtblAdviceNotes & " "

Const lconOrdUndSQL = "SELECT CustNum, OrderNum, " & _
    "Trim(Denom) & Format(Underpayment,'0.00') as Underpay, " & _
    " BankRepPrintDate, RefundReason as Reason, RefundOrignNum as OrignNum" & _
    " FROM " & gtblAdviceNotes & " "

Const lconRefChqSQL = "SELECT CustNum, OrderNum, Trim(Denom) & Format(TotalIncVat,'0.00') as Amount, " & _
    "ChequeNum, ChequePrintedDate as PrintedDate, ChequePrintedBy as PrintedBy, " & _
    "ChequeClearedDate as ClearedDate, RefundReason as Reason, RefundOrignNum as OrignNum, CardName, " & _
    " format(ChequeRequestDate,'dd/mm/yy') as RequestDate FROM " & gtblAdviceNotes & " "

Const lconRefCCSQL = "SELECT CustNum, OrderNum, Trim(Denom) & Format(IIf(Trim$([OrderType])='C',[Payment],0)" & _
    "+IIf(Trim$([PaymentType2])='C',[Payment2],0),'0.00') AS Amount, BankRepPrintDate, CardNumber, CardName, CardStartDate, " & _
    "CardIssueNumber as IssueNum, CardType, AuthorisationCode as AuthCode FROM " & gtblAdviceNotes & " "
    
Dim lstrFixedSQL As String

    dbgCashbook.ClearFields
    
    If lstrLastTabItem <> pstrKey Then
        chkGenericChk.Value = vbUnchecked
    End If
    
    If lstrExtraSQL = "WHERE 1=0;" Then
        lstrWhereClause = ""
    End If
    
    Select Case pstrKey
    Case "OCC" 'Order(s) Credit Card payments
        chkGenericChk.Caption = "Show only claimed items"
        lblDelete.Visible = False
        If lstrExtraSQL <> "WHERE 1=0;" Then
            lstrWhereClause = " (OrderStatus='D' Or OrderStatus='E' Or OrderStatus='B' Or " & _
                "OrderStatus='C') AND ((IIf(Trim$([OrderType])='C','C','')" & _
                    "+IIf(Trim$([PaymentType2])='C','C',''))='C') "
            If chkGenericChk.Value = vbChecked Then
                lstrWhereClause = lstrWhereClause & "  AND " & _
                    "(BankRepPrintDate Is Not Null And BankRepPrintDate<>#12/30/1899#);"
            ElseIf chkGenericChk.Value = vbUnchecked Then
                lstrWhereClause = lstrWhereClause & "" ' AND (RefundReason<>'' And RefundReason Is Not Null);"
            End If
        End If
        lstrFixedSQL = lconOrdCCSQL & lstrExtraSQL & lstrWhereClause

    Case "OUP" 'Order(s) Under payments
        chkGenericChk.Caption = "Show only paid items"
        lblDelete.Visible = True
        lblDelete = "To remove an underpayment entry, enter Bank Rep Print Date of 31/12/1999"
        If lstrExtraSQL <> "WHERE 1=0;" Then
            lstrWhereClause = " (RefundReason='UNDERPAY' ) "
            If chkGenericChk.Value = vbChecked Then
                lstrWhereClause = lstrWhereClause & "  AND " & _
                    "(BankRepPrintDate Is Not Null And BankRepPrintDate<>#12/30/1899#);"
            ElseIf chkGenericChk.Value = vbUnchecked Then
                lstrWhereClause = lstrWhereClause & " AND (RefundReason<>'' And RefundReason Is Not Null);"
            End If
        End If
        lstrFixedSQL = lconOrdUndSQL & lstrExtraSQL & lstrWhereClause

    Case "RCQ" 'Refund(s) Cheque payments
        chkGenericChk.Caption = "Show only cleared items"
        lblDelete.Visible = True
        lblDelete = "To remove a cheque entry enter Bank Rep Print Date of 31/12/1999"
        If lstrExtraSQL <> "WHERE 1=0;" Then
            lstrWhereClause = " (RefundReason='OVERPAY' Or RefundReason='STOCKOUT' " & _
                "Or RefundReason='REFUND' Or RefundReason='OUTOFSTOCK') "
            If chkGenericChk.Value = vbChecked Then
                lstrWhereClause = lstrWhereClause & "  AND " & _
                    "(ChequeClearedDate Is Not Null And ChequeClearedDate<>#12/30/1899#);"
            ElseIf chkGenericChk.Value = vbUnchecked Then
                lstrWhereClause = lstrWhereClause & " AND (RefundReason<>'' And RefundReason Is Not Null);"
            End If
        End If
        lstrFixedSQL = lconRefChqSQL & lstrExtraSQL & lstrWhereClause
                
    Case "RCC" 'Refund(s) Credit Card Payments
        chkGenericChk.Caption = "Show only authorised items"
        lblDelete.Visible = False
        If lstrExtraSQL <> "WHERE 1=0;" Then
            lstrWhereClause = " (OrderStatus='R') AND ((IIf(Trim$([OrderType])='C','C','')" & _
                    "+IIf(Trim$([PaymentType2])='C','C',''))='C') "
            If chkGenericChk.Value = vbChecked Then
                lstrWhereClause = lstrWhereClause & "  AND " & _
                    "(BankRepPrintDate Is Not Null And BankRepPrintDate<>#12/30/1899#);"
            ElseIf chkGenericChk.Value = vbUnchecked Then
                lstrWhereClause = lstrWhereClause & ""
            End If
        End If
        lstrFixedSQL = lconRefCCSQL & lstrExtraSQL & lstrWhereClause

    End Select
            
    datCashbook.RecordSource = lstrFixedSQL
    datCashbook.Refresh
    dbgCashbook.ReBind

    ResizeColumns

    lstrLastTabItem = pstrKey
    
    lblFoundNumber = "Found " & datCashbook.Recordset.RecordCount & " records."
        
End Sub
Sub ResizeColumns()
Dim lintArrInc As Integer
Dim llngThisColWidth As Long

    For lintArrInc = 0 To dbgCashbook.Columns.Count - 1
        Select Case dbgCashbook.Columns(lintArrInc).Caption
        Case "Amount"
            llngThisColWidth = 700.2363
            dbgCashbook.Columns(lintArrInc).Alignment = 1
        Case "Underpay", "Underpayment"
            llngThisColWidth = 800.2363
            dbgCashbook.Columns(lintArrInc).Alignment = 1
        Case "BankRepPrintDate"
            llngThisColWidth = 1544.882
            dbgCashbook.Columns(lintArrInc).NumberFormat = "short date"
        Case "CardName":            llngThisColWidth = 1544.882
        Case "CardStartDate", "StartDate"
            llngThisColWidth = 1094.74
            dbgCashbook.Columns(lintArrInc).NumberFormat = "short date"
        Case "CardIssueNumber":     llngThisColWidth = 1470.047
        Case "CardType":            llngThisColWidth = 945.0709
        Case "CardNumber":          llngThisColWidth = 1300
        Case "AuthorisationCode":   llngThisColWidth = 1395.213
        Case "CustNum":             llngThisColWidth = 870.2363
        Case "OrderNum":            llngThisColWidth = 870.2363
        Case "ChequeNum":           llngThisColWidth = 1094.74
        Case "ChequePrintedDate", "PrintedDate"
            llngThisColWidth = 1020 '1620.284
            dbgCashbook.Columns(lintArrInc).NumberFormat = "short date"
        Case "ChequePrintedBy":     llngThisColWidth = 1395.213
        Case "Reconcil"
            llngThisColWidth = 870.2363
            dbgCashbook.Columns(lintArrInc).Alignment = 1
        Case "ChequeClearedDate", "ClearedDate"
            llngThisColWidth = 1020 '1620.284
        Case "ChequeRequestDate", "RequestDate"
            llngThisColWidth = 1020 '1620.284
            dbgCashbook.Columns(lintArrInc).NumberFormat = "short date"
        Case "RefundReason", "Reason":       llngThisColWidth = 1200.976
        Case "RefundOrignNum", "OrignNum":     llngThisColWidth = 870.2363
        End Select
            
        dbgCashbook.Columns(lintArrInc).Width = llngThisColWidth
    Next lintArrInc
    
End Sub
