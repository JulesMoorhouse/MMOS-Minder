VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildCashbook 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customer Cash Book"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10650
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkBankRepPrintDate 
      Caption         =   "Only show processed cheques "
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   4470
      Width           =   375
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   8760
      Top             =   360
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4470
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
      Top             =   480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1305
   End
   Begin VB.Frame fraSearchBy 
      Caption         =   "Search By"
      Height          =   855
      Left            =   4560
      TabIndex        =   10
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton optSearchField 
         Caption         =   "C&heque Number"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Tag             =   "BPCS"
         Top             =   240
         Value           =   -1  'True
         Width           =   1812
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
         Caption         =   "&Customer Number"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   "CustNumber"
         Top             =   480
         Width           =   1692
      End
   End
   Begin VB.TextBox txtSearchCriteria 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2412
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4470
      Width           =   1305
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   14
      Top             =   4920
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13600
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "09/06/2019"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "13:05"
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid dbgCashbook 
      Bindings        =   "ChdCshbk.frx":0000
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "ChdCshbk.frx":001A
      TabIndex        =   15
      Top             =   1080
      Width           =   10335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Criteria :-"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   "To remove a cheque entry, enter Bank Rep Print Date of 31/12/1999"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   4200
      Width           =   5535
   End
   Begin VB.Label lblFoundNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Found 0 records"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   2415
   End
End
Attribute VB_Name = "frmChildCashbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrRoute As String
Dim lstrScreenHelpFile As String

Const lconFixedSQL = "SELECT ChequeNum, ChequePrintedDate, ChequePrintedBy, " & _
        "Trim(Denom) & Format(Reconcilliation,'0.00') as Reconcil, " & _
        "Trim(Denom) & Format(Underpayment,'0.00') as Underpay, " & _
        "ChequeClearedDate, RefundReason, RefundOrignNum, CardName, " & _
        "BankRepPrintDate, CustNum, OrderNum, ChequeRequestDate FROM " & gtblAdviceNotes & " " '& _
        " WHERE (((RefundReason)<>'' And (RefundReason) Is Not Null) "

Dim lstrExtraSQL As String
Dim lstrWhereClause As String
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property

Private Sub chkBankRepPrintDate_Click()

    If chkBankRepPrintDate.Value = vbChecked Then
        lstrWhereClause = " ((RefundReason<>'' And RefundReason Is Not Null) AND " & _
            "(BankRepPrintDate Is Not Null And BankRepPrintDate<>#12/30/1899#));"
                    
    ElseIf chkBankRepPrintDate.Value = vbUnchecked Then
        lstrWhereClause = " (RefundReason<>'' And RefundReason Is Not Null);"

    End If
    datCashbook.RecordSource = lconFixedSQL & lstrExtraSQL & lstrWhereClause
    datCashbook.Refresh
    
    lblFoundNumber = "Found " & datCashbook.Recordset.RecordCount & " records."
    
End Sub

Private Sub cmdBack_Click()

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

    lstrExtraSQL = "where CustNum = " & gstrCustomerAccount.lngCustNum & " and "

    If optSearchField(0).Value = True Then  ' Cheque Number
        If (txtSearchCriteria) <> "" Then
            lstrExtraSQL = "where ChequeNum = '" & Trim$(txtSearchCriteria) & "' and "
        End If
    ElseIf optSearchField(1).Value = True Then ' Customer Number
        If CLng(Val(txtSearchCriteria)) = 0 Then
            MsgBox "You must enter a Customer Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where CustNum = " & CLng(txtSearchCriteria) & " and "
        End If
    ElseIf optSearchField(2).Value = True Then ' Order Number
        If CLng(Val(txtSearchCriteria)) = 0 Then
            MsgBox "You must enter a Order Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where CustNum = " & CLng(Val(txtSearchCriteria)) & " and "
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

    datCashbook.RecordSource = lconFixedSQL & lstrExtraSQL & lstrWhereClause
            
    datCashbook.Refresh
    dbgCashbook.Refresh
    
    lblFoundNumber = "Found " & datCashbook.Recordset.RecordCount & " records."

    Busy False, Me

End Sub

Private Sub cmdHelp_Click()

    RunNDontWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1009.htm"

End Sub

Private Sub Form_Activate()

    lblFoundNumber = "Found " & datCashbook.Recordset.RecordCount & " records."
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    On Error GoTo ErrHandler
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datCashbook.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datCashbook.DatabaseName = gstrStatic.strCentralDBFile
    End Select
    
    If gstrSystemRoute <> srCompanyRoute Then
        datCashbook.Connect = gstrDBPasswords.strCentralDBPasswordString
    End If
    
    Select Case Me.Route
    Case gconstrCashbookSpecificCustomer
        datCashbook.RecordSource = "SELECT ChequeNum, ChequePrintedDate, ChequePrintedBy, " & _
            "Trim(Denom) & Format(Reconcilliation,'0.00') as Reconcil, " & _
            "Trim(Denom) & Format(Underpayment,'0.00') as Underpay, " & _
            "ChequeClearedDate, RefundReason, RefundOrignNum, CardName, " & _
            "BankRepPrintDate, CustNum, OrderNum, ChequeRequestDate FROM " & gtblAdviceNotes & _
            " WHERE (((RefundReason)<>'' And (RefundReason) Is Not Null) " & _
            "AND ((BankRepPrintDate) Is Not Null And (AdviceNotes.BankRepPrintDate)<>#00:00:00#) " & _
            "AND ((AdviceNotes.CustNum)=" & gstrCustomerAccount.lngCustNum & "));"
            
        dbgCashbook.AllowDelete = False
        lblDelete.Visible = True
        optSearchField(1).Enabled = False
    Case Else
        MsgBox "Error please report where this happened!", vbInformation, gconstrTitlPrefix & "Child Cash Book"
        
    End Select
    
    datCashbook.Refresh
    
    lstrExtraSQL = "where CustNum = " & gstrCustomerAccount.lngCustNum & " and "

    lstrWhereClause = " ((RefundReason<>'' And RefundReason Is Not Null) AND " & _
        "(BankRepPrintDate Is Not Null And BankRepPrintDate<>#12/30/1899#));"
    
    GetLocalFields
    
    SetupHelpFileReqs

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmChildCashBook.Load", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
            
    
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
